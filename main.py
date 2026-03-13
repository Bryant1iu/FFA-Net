#!/usr/bin/env python3
"""
FFA-Net: Feature Fusion Attention Network for Single Image Dehazing
Unified entry point — supports training and inference via subcommands.

Usage:
    python main.py train --data_dir ./data --trainset its --steps 500000
    python main.py test  --model_path trained_models/its_train_ffa_3_19.pk --input test_imgs --output results
"""

import argparse
import math
import os
import random
import sys
import time
import warnings

import numpy as np
import torch
import torch.nn as nn
import torch.nn.functional as F
import torch.utils.data as data
import torchvision.transforms as tfs
import torchvision.utils as vutils
from PIL import Image
from torch import optim
from torch.backends import cudnn
from torch.utils.data import DataLoader
from torchvision.transforms import functional as TF

warnings.filterwarnings("ignore")

# ═══════════════════════════════════════════════════════════
# Model Architecture
# ═══════════════════════════════════════════════════════════

def default_conv(in_channels, out_channels, kernel_size, bias=True):
    return nn.Conv2d(in_channels, out_channels, kernel_size,
                     padding=(kernel_size // 2), bias=bias)


class PALayer(nn.Module):
    """Pixel (Spatial) Attention Layer."""
    def __init__(self, channel):
        super().__init__()
        self.pa = nn.Sequential(
            nn.Conv2d(channel, channel // 8, 1, padding=0, bias=True),
            nn.ReLU(inplace=True),
            nn.Conv2d(channel // 8, 1, 1, padding=0, bias=True),
            nn.Sigmoid(),
        )

    def forward(self, x):
        return x * self.pa(x)


class CALayer(nn.Module):
    """Channel Attention Layer."""
    def __init__(self, channel):
        super().__init__()
        self.avg_pool = nn.AdaptiveAvgPool2d(1)
        self.ca = nn.Sequential(
            nn.Conv2d(channel, channel // 8, 1, padding=0, bias=True),
            nn.ReLU(inplace=True),
            nn.Conv2d(channel // 8, channel, 1, padding=0, bias=True),
            nn.Sigmoid(),
        )

    def forward(self, x):
        return x * self.ca(self.avg_pool(x))


class Block(nn.Module):
    """Residual block with dual attention (CA + PA)."""
    def __init__(self, conv, dim, kernel_size):
        super().__init__()
        self.conv1 = conv(dim, dim, kernel_size, bias=True)
        self.act1 = nn.ReLU(inplace=True)
        self.conv2 = conv(dim, dim, kernel_size, bias=True)
        self.calayer = CALayer(dim)
        self.palayer = PALayer(dim)

    def forward(self, x):
        res = self.act1(self.conv1(x))
        res = res + x
        res = self.conv2(res)
        res = self.calayer(res)
        res = self.palayer(res)
        res += x
        return res


class Group(nn.Module):
    """A group of residual blocks with skip connection."""
    def __init__(self, conv, dim, kernel_size, blocks):
        super().__init__()
        modules = [Block(conv, dim, kernel_size) for _ in range(blocks)]
        modules.append(conv(dim, dim, kernel_size))
        self.gp = nn.Sequential(*modules)

    def forward(self, x):
        return self.gp(x) + x


class FFA(nn.Module):
    """Feature Fusion Attention Network."""
    def __init__(self, gps=3, blocks=19, conv=default_conv):
        super().__init__()
        self.gps = gps
        self.dim = 64
        kernel_size = 3
        assert self.gps == 3

        self.pre = nn.Sequential(conv(3, self.dim, kernel_size))
        self.g1 = Group(conv, self.dim, kernel_size, blocks=blocks)
        self.g2 = Group(conv, self.dim, kernel_size, blocks=blocks)
        self.g3 = Group(conv, self.dim, kernel_size, blocks=blocks)
        self.ca = nn.Sequential(
            nn.AdaptiveAvgPool2d(1),
            nn.Conv2d(self.dim * self.gps, self.dim // 16, 1, padding=0),
            nn.ReLU(inplace=True),
            nn.Conv2d(self.dim // 16, self.dim * self.gps, 1, padding=0, bias=True),
            nn.Sigmoid(),
        )
        self.palayer = PALayer(self.dim)
        self.post = nn.Sequential(
            conv(self.dim, self.dim, kernel_size),
            conv(self.dim, 3, kernel_size),
        )

    def forward(self, x1):
        x = self.pre(x1)
        res1 = self.g1(x)
        res2 = self.g2(res1)
        res3 = self.g3(res2)
        w = self.ca(torch.cat([res1, res2, res3], dim=1))
        w = w.view(-1, self.gps, self.dim)[:, :, :, None, None]
        out = w[:, 0, ::] * res1 + w[:, 1, ::] * res2 + w[:, 2, ::] * res3
        out = self.palayer(out)
        x = self.post(out)
        return x + x1


class LossNetwork(nn.Module):
    """VGG16-based perceptual loss."""
    def __init__(self, vgg_model):
        super().__init__()
        self.vgg_layers = vgg_model
        self.layer_name_mapping = {"3": "relu1_2", "8": "relu2_2", "15": "relu3_3"}

    def output_features(self, x):
        output = {}
        for name, module in self.vgg_layers._modules.items():
            x = module(x)
            if name in self.layer_name_mapping:
                output[self.layer_name_mapping[name]] = x
        return list(output.values())

    def forward(self, dehaze, gt):
        loss = []
        dehaze_features = self.output_features(dehaze)
        gt_features = self.output_features(gt)
        for df, gf in zip(dehaze_features, gt_features):
            loss.append(F.mse_loss(df, gf))
        return sum(loss) / len(loss)


# ═══════════════════════════════════════════════════════════
# Metrics
# ═══════════════════════════════════════════════════════════

def _gaussian(window_size, sigma):
    gauss = torch.Tensor(
        [math.exp(-(x - window_size // 2) ** 2 / (2 * sigma ** 2)) for x in range(window_size)]
    )
    return gauss / gauss.sum()


def _create_window(window_size, channel):
    _1d = _gaussian(window_size, 1.5).unsqueeze(1)
    _2d = _1d.mm(_1d.t()).float().unsqueeze(0).unsqueeze(0)
    return _2d.expand(channel, 1, window_size, window_size).contiguous()


def ssim(img1, img2, window_size=11, size_average=True):
    img1 = torch.clamp(img1, 0, 1)
    img2 = torch.clamp(img2, 0, 1)
    channel = img1.size(1)
    window = _create_window(window_size, channel)
    if img1.is_cuda:
        window = window.cuda(img1.get_device())
    window = window.type_as(img1)

    mu1 = F.conv2d(img1, window, padding=window_size // 2, groups=channel)
    mu2 = F.conv2d(img2, window, padding=window_size // 2, groups=channel)
    mu1_sq, mu2_sq, mu1_mu2 = mu1.pow(2), mu2.pow(2), mu1 * mu2
    sigma1_sq = F.conv2d(img1 * img1, window, padding=window_size // 2, groups=channel) - mu1_sq
    sigma2_sq = F.conv2d(img2 * img2, window, padding=window_size // 2, groups=channel) - mu2_sq
    sigma12 = F.conv2d(img1 * img2, window, padding=window_size // 2, groups=channel) - mu1_mu2
    C1, C2 = 0.01 ** 2, 0.03 ** 2
    ssim_map = ((2 * mu1_mu2 + C1) * (2 * sigma12 + C2)) / (
        (mu1_sq + mu2_sq + C1) * (sigma1_sq + sigma2_sq + C2)
    )
    return ssim_map.mean() if size_average else ssim_map.mean(1).mean(1).mean(1)


def psnr(pred, gt):
    pred = pred.clamp(0, 1).detach().cpu().numpy()
    gt = gt.clamp(0, 1).detach().cpu().numpy()
    rmse = math.sqrt(np.mean((pred - gt) ** 2))
    if rmse == 0:
        return 100
    return 20 * math.log10(1.0 / rmse)


# ═══════════════════════════════════════════════════════════
# Dataset
# ═══════════════════════════════════════════════════════════

IMG_NORM_MEAN = [0.64, 0.6, 0.58]
IMG_NORM_STD = [0.14, 0.15, 0.152]


class RESIDEDataset(data.Dataset):
    def __init__(self, path, train, size="whole_img", fmt=".png"):
        super().__init__()
        self.size = size
        self.train = train
        self.fmt = fmt
        self.haze_imgs = [
            os.path.join(path, "hazy", f) for f in os.listdir(os.path.join(path, "hazy"))
        ]
        self.clear_dir = os.path.join(path, "clear")

    def __getitem__(self, index):
        haze = Image.open(self.haze_imgs[index])
        if isinstance(self.size, int):
            while haze.size[0] < self.size or haze.size[1] < self.size:
                index = random.randint(0, len(self.haze_imgs) - 1)
                haze = Image.open(self.haze_imgs[index])
        img_id = os.path.basename(self.haze_imgs[index]).split("_")[0]
        clear = Image.open(os.path.join(self.clear_dir, img_id + self.fmt))
        clear = tfs.CenterCrop(haze.size[::-1])(clear)
        if isinstance(self.size, int):
            i, j, h, w = tfs.RandomCrop.get_params(haze, output_size=(self.size, self.size))
            haze = TF.crop(haze, i, j, h, w)
            clear = TF.crop(clear, i, j, h, w)
        haze, clear = self._augment(haze.convert("RGB"), clear.convert("RGB"))
        return haze, clear

    def _augment(self, data_img, target):
        if self.train:
            rand_hor = random.randint(0, 1)
            rand_rot = random.randint(0, 3)
            data_img = tfs.RandomHorizontalFlip(rand_hor)(data_img)
            target = tfs.RandomHorizontalFlip(rand_hor)(target)
            if rand_rot:
                data_img = TF.rotate(data_img, 90 * rand_rot)
                target = TF.rotate(target, 90 * rand_rot)
        data_img = tfs.Normalize(mean=IMG_NORM_MEAN, std=IMG_NORM_STD)(tfs.ToTensor()(data_img))
        target = tfs.ToTensor()(target)
        return data_img, target

    def __len__(self):
        return len(self.haze_imgs)


def build_dataloaders(data_dir, trainset, testset, batch_size, crop, crop_size):
    """Build train and test dataloaders based on dataset names."""
    size = crop_size if crop else "whole_img"
    reside = os.path.join(data_dir, "RESIDE")

    loader_map = {
        "its_train": lambda: DataLoader(
            RESIDEDataset(os.path.join(reside, "ITS"), train=True, size=size, fmt=".png"),
            batch_size=batch_size, shuffle=True,
        ),
        "its_test": lambda: DataLoader(
            RESIDEDataset(os.path.join(reside, "SOTS", "indoor"), train=False, size="whole_img", fmt=".png"),
            batch_size=1, shuffle=False,
        ),
        "ots_train": lambda: DataLoader(
            RESIDEDataset(os.path.join(reside, "OTS"), train=True, size=size, fmt=".jpg"),
            batch_size=batch_size, shuffle=True,
        ),
        "ots_test": lambda: DataLoader(
            RESIDEDataset(os.path.join(reside, "SOTS", "outdoor"), train=False, size="whole_img", fmt=".png"),
            batch_size=1, shuffle=False,
        ),
    }
    train_loader = loader_map[trainset]()
    test_loader = loader_map[testset]()
    return train_loader, test_loader


# ═══════════════════════════════════════════════════════════
# Training
# ═══════════════════════════════════════════════════════════

def lr_schedule_cosdecay(t, T, init_lr):
    return 0.5 * (1 + math.cos(t * math.pi / T)) * init_lr


def evaluate(net, loader_test, device):
    """Run evaluation on test set, return (mean_ssim, mean_psnr)."""
    net.eval()
    torch.cuda.empty_cache()
    ssims, psnrs = [], []
    for inputs, targets in loader_test:
        inputs = inputs.to(device)
        targets = targets.to(device)
        pred = net(inputs)
        ssims.append(ssim(pred, targets).item())
        psnrs.append(psnr(pred, targets))
    return np.mean(ssims), np.mean(psnrs)


def do_train(args):
    """Execute the training pipeline."""
    device = "cuda" if torch.cuda.is_available() else "cpu"
    model_name = f"{args.trainset}_ffa_{args.gps}_{args.blocks}"

    # Directories
    os.makedirs(args.model_dir, exist_ok=True)
    os.makedirs("numpy_files", exist_ok=True)
    os.makedirs(f"samples/{model_name}", exist_ok=True)
    os.makedirs(f"logs/{model_name}", exist_ok=True)

    model_path = os.path.join(args.model_dir, model_name + ".pk")
    print(f"Model checkpoint: {model_path}")
    print(f"Device: {device}")

    # Model
    net = FFA(gps=args.gps, blocks=args.blocks).to(device)
    if device == "cuda":
        net = nn.DataParallel(net)
        cudnn.benchmark = True

    # Data
    loader_train, loader_test = build_dataloaders(
        args.data_dir, args.trainset, args.testset,
        args.bs, args.crop, args.crop_size,
    )

    # Loss
    criterion = [nn.L1Loss().to(device)]
    if args.perloss:
        from torchvision.models import vgg16
        vgg_model = vgg16(pretrained=True).features[:16].to(device)
        for p in vgg_model.parameters():
            p.requires_grad = False
        criterion.append(LossNetwork(vgg_model).to(device))

    optimizer = optim.Adam(
        filter(lambda x: x.requires_grad, net.parameters()),
        lr=args.lr, betas=(0.9, 0.999), eps=1e-08,
    )
    optimizer.zero_grad()

    # Resume
    losses, ssims_hist, psnrs_hist = [], [], []
    start_step, max_ssim, max_psnr = 0, 0, 0
    if args.resume and os.path.exists(model_path):
        print(f"Resuming from {model_path}")
        ckp = torch.load(model_path, map_location=device)
        net.load_state_dict(ckp["model"])
        losses = ckp["losses"]
        start_step = ckp["step"]
        max_ssim = ckp["max_ssim"]
        max_psnr = ckp["max_psnr"]
        psnrs_hist = ckp["psnrs"]
        ssims_hist = ckp["ssims"]
        print(f"Resumed at step {start_step}")
    else:
        print("Training from scratch")

    start_time = time.time()
    T = args.steps

    for step in range(start_step + 1, T + 1):
        net.train()
        lr = args.lr
        if not args.no_lr_sche:
            lr = lr_schedule_cosdecay(step, T, args.lr)
            for pg in optimizer.param_groups:
                pg["lr"] = lr

        x, y = next(iter(loader_train))
        x, y = x.to(device), y.to(device)
        out = net(x)
        loss = criterion[0](out, y)
        if args.perloss:
            loss = loss + 0.04 * criterion[1](out, y)

        loss.backward()
        optimizer.step()
        optimizer.zero_grad()
        losses.append(loss.item())
        elapsed = (time.time() - start_time) / 60
        print(
            f"\rtrain loss: {loss.item():.5f} | step: {step}/{T} | "
            f"lr: {lr:.7f} | time: {elapsed:.1f}min",
            end="", flush=True,
        )

        if step % args.eval_step == 0:
            with torch.no_grad():
                ssim_eval, psnr_eval = evaluate(net, loader_test, device)
            print(f"\nstep: {step} | ssim: {ssim_eval:.4f} | psnr: {psnr_eval:.4f}")
            ssims_hist.append(ssim_eval)
            psnrs_hist.append(psnr_eval)
            if ssim_eval > max_ssim and psnr_eval > max_psnr:
                max_ssim = ssim_eval
                max_psnr = psnr_eval
                torch.save(
                    {
                        "step": step,
                        "max_psnr": max_psnr,
                        "max_ssim": max_ssim,
                        "ssims": ssims_hist,
                        "psnrs": psnrs_hist,
                        "losses": losses,
                        "model": net.state_dict(),
                    },
                    model_path,
                )
                print(f"Model saved | max_psnr: {max_psnr:.4f} | max_ssim: {max_ssim:.4f}")

    np.save(f"numpy_files/{model_name}_{T}_losses.npy", losses)
    np.save(f"numpy_files/{model_name}_{T}_ssims.npy", ssims_hist)
    np.save(f"numpy_files/{model_name}_{T}_psnrs.npy", psnrs_hist)
    print("\nTraining finished.")


# ═══════════════════════════════════════════════════════════
# Inference
# ═══════════════════════════════════════════════════════════

def do_test(args):
    """Run inference on images in a folder."""
    device = "cuda" if torch.cuda.is_available() else "cpu"
    os.makedirs(args.output, exist_ok=True)
    print(f"Device: {device}")
    print(f"Model : {args.model_path}")
    print(f"Input : {args.input}")
    print(f"Output: {args.output}")

    ckp = torch.load(args.model_path, map_location=device)
    net = FFA(gps=args.gps, blocks=args.blocks)
    net = nn.DataParallel(net)
    net.load_state_dict(ckp["model"])
    net = net.to(device)
    net.eval()

    normalize = tfs.Compose([
        tfs.ToTensor(),
        tfs.Normalize(mean=IMG_NORM_MEAN, std=IMG_NORM_STD),
    ])

    img_files = [f for f in os.listdir(args.input) if f.lower().endswith((".png", ".jpg", ".jpeg", ".bmp"))]
    print(f"Found {len(img_files)} images")

    for im_name in img_files:
        haze = Image.open(os.path.join(args.input, im_name)).convert("RGB")
        haze_tensor = normalize(haze).unsqueeze(0).to(device)
        with torch.no_grad():
            pred = net(haze_tensor)
        pred = torch.squeeze(pred.clamp(0, 1).cpu())
        out_name = os.path.splitext(im_name)[0] + "_FFA.png"
        vutils.save_image(pred, os.path.join(args.output, out_name))
        print(f"\r  {im_name} -> {out_name}", end="", flush=True)

    print(f"\nDone. {len(img_files)} images saved to {args.output}")


# ═══════════════════════════════════════════════════════════
# CLI
# ═══════════════════════════════════════════════════════════

def main():
    parser = argparse.ArgumentParser(
        description="FFA-Net: Feature Fusion Attention Network for Single Image Dehazing",
    )
    subparsers = parser.add_subparsers(dest="command", help="Available commands")

    # ── train ──
    p_train = subparsers.add_parser("train", help="Train FFA-Net")
    p_train.add_argument("--data_dir", type=str, required=True, help="Root data directory (contains RESIDE/)")
    p_train.add_argument("--trainset", type=str, default="its_train", choices=["its_train", "ots_train"])
    p_train.add_argument("--testset", type=str, default="its_test", choices=["its_test", "ots_test"])
    p_train.add_argument("--steps", type=int, default=500000, help="Total training steps")
    p_train.add_argument("--eval_step", type=int, default=5000, help="Evaluate every N steps")
    p_train.add_argument("--lr", type=float, default=1e-4, help="Learning rate")
    p_train.add_argument("--bs", type=int, default=2, help="Batch size")
    p_train.add_argument("--gps", type=int, default=3, help="Number of residual groups")
    p_train.add_argument("--blocks", type=int, default=19, help="Residual blocks per group")
    p_train.add_argument("--crop", action="store_true", help="Enable random cropping")
    p_train.add_argument("--crop_size", type=int, default=240, help="Crop size (requires --crop)")
    p_train.add_argument("--no_lr_sche", action="store_true", help="Disable cosine LR schedule")
    p_train.add_argument("--perloss", action="store_true", help="Enable VGG16 perceptual loss")
    p_train.add_argument("--resume", action="store_true", default=True, help="Resume from checkpoint")
    p_train.add_argument("--model_dir", type=str, default="trained_models", help="Checkpoint directory")

    # ── test ──
    p_test = subparsers.add_parser("test", help="Run inference on hazy images")
    p_test.add_argument("--model_path", type=str, required=True, help="Path to .pk checkpoint")
    p_test.add_argument("--input", type=str, required=True, help="Directory of input hazy images")
    p_test.add_argument("--output", type=str, default="results", help="Directory for dehazed outputs")
    p_test.add_argument("--gps", type=int, default=3, help="Number of residual groups")
    p_test.add_argument("--blocks", type=int, default=19, help="Residual blocks per group")

    args = parser.parse_args()
    if args.command is None:
        parser.print_help()
        sys.exit(1)

    if args.command == "train":
        do_train(args)
    elif args.command == "test":
        do_test(args)


if __name__ == "__main__":
    main()
