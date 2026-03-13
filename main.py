#!/usr/bin/env python3
"""
FFA-Net: Feature Fusion Attention Network for Single Image Dehazing

Usage:
    python main.py train --data_dir ./data --trainset its_train --steps 500000
    python main.py test  --model_path trained_models/its_train_ffa_3_19.pk --input test_imgs --output results
"""

import argparse
import sys

import ffa_net


def build_parser():
    parser = argparse.ArgumentParser(
        description="FFA-Net: Feature Fusion Attention Network for Single Image Dehazing",
    )
    sub = parser.add_subparsers(dest="command", help="Available commands")

    # ── train ──────────────────────────────────────────────
    p_train = sub.add_parser("train", help="Train FFA-Net on RESIDE dataset")
    p_train.add_argument("--data_dir",  type=str, required=True,
                         help="Root data directory (must contain RESIDE/)")
    p_train.add_argument("--trainset",  type=str, default="its_train",
                         choices=["its_train", "ots_train"])
    p_train.add_argument("--testset",   type=str, default="its_test",
                         choices=["its_test", "ots_test"])
    p_train.add_argument("--steps",     type=int, default=500000,
                         help="Total training steps")
    p_train.add_argument("--eval_step", type=int, default=5000,
                         help="Evaluate every N steps")
    p_train.add_argument("--lr",        type=float, default=1e-4,
                         help="Initial learning rate")
    p_train.add_argument("--bs",        type=int, default=2,
                         help="Batch size")
    p_train.add_argument("--gps",       type=int, default=3,
                         help="Number of residual groups (must be 3)")
    p_train.add_argument("--blocks",    type=int, default=19,
                         help="Residual blocks per group")
    p_train.add_argument("--crop",      action="store_true",
                         help="Enable random cropping during training")
    p_train.add_argument("--crop_size", type=int, default=240,
                         help="Crop size (requires --crop)")
    p_train.add_argument("--no_lr_sche", action="store_true",
                         help="Disable cosine LR decay schedule")
    p_train.add_argument("--perloss",   action="store_true",
                         help="Enable VGG16 perceptual loss")
    p_train.add_argument("--resume",    action="store_true", default=True,
                         help="Resume training from existing checkpoint")
    p_train.add_argument("--model_dir", type=str, default="trained_models",
                         help="Directory for saving checkpoints")

    # ── test ───────────────────────────────────────────────
    p_test = sub.add_parser("test", help="Run inference on hazy images")
    p_test.add_argument("--model_path", type=str, required=True,
                        help="Path to trained .pk checkpoint")
    p_test.add_argument("--input",      type=str, required=True,
                        help="Directory of input hazy images")
    p_test.add_argument("--output",     type=str, default="results",
                        help="Directory for dehazed output images")
    p_test.add_argument("--gps",        type=int, default=3,
                        help="Number of residual groups (must match checkpoint)")
    p_test.add_argument("--blocks",     type=int, default=19,
                        help="Residual blocks per group (must match checkpoint)")

    return parser


def main():
    parser = build_parser()
    args = parser.parse_args()

    if args.command is None:
        parser.print_help()
        sys.exit(1)

    if args.command == "train":
        ffa_net.train(args)
    elif args.command == "test":
        ffa_net.test(args)


if __name__ == "__main__":
    main()
