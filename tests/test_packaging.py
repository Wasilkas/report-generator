import os
import subprocess
import sys
import zipfile
from pathlib import Path


def test_wheel_resources_installed_entrypoint(tmp_path):
    root = Path(__file__).parents[1]
    subprocess.run(
        ['uv', 'build', '--wheel', '--sdist', '--out-dir', str(tmp_path), str(root)],
        check=True,
        capture_output=True,
        text=True,
    )
    wheel = next(tmp_path.glob('*.whl'))
    with zipfile.ZipFile(wheel) as archive:
        assert 'report_generator/configs/config.yaml' in archive.namelist()
        assert 'report_generator/configs/app_config.yaml' in archive.namelist()
    installed = tmp_path / 'installed'
    subprocess.run(
        [
            'uv',
            'pip',
            'install',
            '--python',
            sys.executable,
            '--no-deps',
            '--target',
            str(installed),
            str(wheel),
        ],
        check=True,
        capture_output=True,
        text=True,
    )
    env = os.environ.copy()
    env['PYTHONPATH'] = str(installed)
    result = subprocess.run(
        [
            sys.executable,
            '-c',
            'from report_generator.config import Config; print(Config.default().min_train_count)',
        ],
        cwd=tmp_path,
        env=env,
        check=True,
        capture_output=True,
        text=True,
    )
    assert result.stdout.strip() == '20'
    result = subprocess.run(
        [str(installed / 'bin' / 'generate-report'), '--help'],
        cwd=tmp_path,
        env=env,
        check=True,
        capture_output=True,
        text=True,
    )
    assert 'business' in result.stdout and 'dev' in result.stdout
