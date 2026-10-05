"""Generate two dummy metrics Excel files for report testing."""

import argparse
import os
import random
from pathlib import Path

import pandas as pd

CLASSES = [
    'cat',
    'dog',
    'car',
    'truck',
    'bicycle',
    'person',
    'bird',
    'bottle',
    'chair',
    'sofa',
]
# These will be excluded (train count < 20)
RARE_CLASSES = ['rare_object', 'tiny_thing']

ALL_CLASSES = CLASSES + RARE_CLASSES


def make_row(cls: str, rare: bool, rng: random.Random, boost: float = 0.0) -> dict:
    train = rng.randint(3, 10) if rare else rng.randint(25, 300)
    base_f1 = round(rng.uniform(0.55, 0.92) + boost, 4)
    base_f1 = min(base_f1, 0.999)
    prec = round(min(base_f1 + rng.uniform(-0.05, 0.08), 0.999), 4)
    rec = round(min(base_f1 + rng.uniform(-0.05, 0.08), 0.999), 4)
    return {
        '': cls,
        'confidence': round(rng.uniform(0.3, 0.7), 4),
        'precision': prec,
        'recall': rec,
        'f1_score': base_f1,
        'ap50': round(min(base_f1 + rng.uniform(0.0, 0.06), 0.999), 4),
        'ap75': round(min(base_f1 - rng.uniform(0.02, 0.12), 0.999), 4),
        'ap50_95': round(min(base_f1 - rng.uniform(0.05, 0.15), 0.999), 4),
        'perebrak': round(rng.uniform(0.05, 0.45), 4),
        'nedobrak': round(rng.uniform(0.05, 0.35), 4),
        'Количество примеров train': train,
        'Количество примеров test': rng.randint(5, 60),
        'Количество примеров val': rng.randint(5, 40),
    }


def generate(output1: str | Path, output2: str | Path, *, overwrite: bool = False) -> None:
    paths = [Path(output1), Path(output2)]
    if paths[0].resolve() == paths[1].resolve() or (
        all(path.exists() for path in paths) and paths[0].samefile(paths[1])
    ):
        raise ValueError('Sample destinations must be distinct')
    resolved = [path.resolve() for path in paths]
    if resolved[0] in resolved[1].parents or resolved[1] in resolved[0].parents:
        raise ValueError('Sample destinations must not be nested')
    for path in paths:
        if path.suffix.lower() != '.xlsx':
            raise ValueError(f'Sample destination must end in .xlsx: {path}')
        if path.exists() and not path.is_file():
            raise ValueError(f'Sample destination is not a file: {path}')
        ancestor = path.parent
        while not ancestor.exists():
            ancestor = ancestor.parent
        if not ancestor.is_dir() or not os.access(ancestor, os.W_OK):
            raise ValueError(f'Sample destination parent is not writable: {path.parent}')
        if path.exists() and not os.access(path, os.W_OK):
            raise ValueError(f'Sample destination is not writable: {path}')
        if path.exists() and not overwrite:
            raise FileExistsError(f'Sample output already exists: {path}; use --force')
    rng = random.Random(42)
    rows1, rows2 = [], []
    for cls in ALL_CLASSES:
        rare = cls in RARE_CLASSES
        rows1.append(make_row(cls, rare, rng, boost=0.03))
        rows2.append(make_row(cls, rare, rng))
    for path, rows in zip(paths, [rows1, rows2]):
        path.parent.mkdir(parents=True, exist_ok=True)
        pd.DataFrame(rows).to_excel(path, index=False)


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('output1', nargs='?', default='test_model1.xlsx')
    parser.add_argument('output2', nargs='?', default='test_model2.xlsx')
    parser.add_argument('--force', action='store_true')
    args = parser.parse_args()
    try:
        generate(args.output1, args.output2, overwrite=args.force)
    except (ValueError, OSError) as exc:
        parser.error(str(exc))
    print(f'Created {args.output1} and {args.output2}')


if __name__ == '__main__':
    main()
