"""Enforce shared-package independence, including static dynamic-import targets."""

import ast
from pathlib import Path

import pytest

PACKAGE = Path(__file__).parents[1] / 'report_generator'
FORBIDDEN = {'clearml_yolo', 'digital_metrics', 'clearml', 'torch', 'ultralytics'}


def check_dependencies(source: str) -> list[str]:
    violations = []
    tree = ast.parse(source)
    for node in ast.walk(tree):
        targets = []
        if isinstance(node, ast.Import):
            targets = [alias.name for alias in node.names]
        elif isinstance(node, ast.ImportFrom) and node.level == 0 and node.module:
            targets = [node.module]
        elif isinstance(node, ast.Call) and node.args:
            func = node.func
            name = (
                func.id
                if isinstance(func, ast.Name)
                else (func.attr if isinstance(func, ast.Attribute) else '')
            )
            if name in {'__import__', 'import_module'}:
                argument = node.args[0]
                if not isinstance(argument, ast.Constant) or not isinstance(argument.value, str):
                    violations.append('Unbounded dynamic import')
                else:
                    targets = [argument.value]
        violations.extend(target for target in targets if target.split('.')[0] in FORBIDDEN)
    return violations


def test_shared_runtime_has_no_application_or_backend_dependency():
    for path in sorted(PACKAGE.rglob('*.py')):
        assert not check_dependencies(path.read_text()), path


@pytest.mark.parametrize(
    'source',
    [
        'import clearml_yolo',
        'from digital_metrics import Metrics',
        'from clearml import Task',
        "__import__('torch')",
        "importlib.import_module('ultralytics.models')",
        'importlib.import_module(variable)',
    ],
)
def test_forbidden_import_negative_case(source):
    assert check_dependencies(source)
