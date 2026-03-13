"""Generate two dummy metrics Excel files for report testing."""

import random

import numpy as np
import pandas as pd

random.seed(42)
np.random.seed(42)

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


def make_row(cls: str, idx: int, rare: bool, boost: float = 0.0) -> dict:
    train = random.randint(3, 10) if rare else random.randint(25, 300)
    base_f1 = round(random.uniform(0.55, 0.92) + boost, 4)
    base_f1 = min(base_f1, 0.999)
    prec = round(min(base_f1 + random.uniform(-0.05, 0.08), 0.999), 4)
    rec = round(min(base_f1 + random.uniform(-0.05, 0.08), 0.999), 4)
    return {
        '': cls,
        'confidence': round(random.uniform(0.3, 0.7), 4),
        'precision': prec,
        'recall': rec,
        'f1_score': base_f1,
        'ap50': round(min(base_f1 + random.uniform(0.0, 0.06), 0.999), 4),
        'ap75': round(min(base_f1 - random.uniform(0.02, 0.12), 0.999), 4),
        'ap50_95': round(min(base_f1 - random.uniform(0.05, 0.15), 0.999), 4),
        'perebrak': round(random.uniform(0.05, 0.45), 4),
        'nedobrak': round(random.uniform(0.05, 0.35), 4),
        'Количество примеров train': train,
        'Количество примеров test': random.randint(5, 60),
        'Количество примеров val': random.randint(5, 40),
    }


rows1, rows2 = [], []
for i, cls in enumerate(ALL_CLASSES):
    rare = cls in RARE_CLASSES
    # model1 (new) gets a small boost on most metrics
    rows1.append(make_row(cls, i, rare, boost=0.03))
    rows2.append(make_row(cls, i, rare, boost=0.0))

df1 = pd.DataFrame(rows1)
df2 = pd.DataFrame(rows2)

df1.to_excel('test_model1.xlsx', index=False)
df2.to_excel('test_model2.xlsx', index=False)
print('Created test_model1.xlsx and test_model2.xlsx')
