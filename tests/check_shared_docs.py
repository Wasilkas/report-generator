"""Check task-owned Markdown structure, relative links and Python examples."""

import re
from pathlib import Path

ROOT = Path(__file__).parents[1]
PATHS = [ROOT / 'README.md', *sorted((ROOT / 'docs/plans/shared-extraction').glob('*.md'))]


def main() -> None:
    for path in PATHS:
        text = path.read_text()
        assert text.endswith('\n'), path
        assert not re.search(r'[ \t]+$', text, re.MULTILINE), path
        assert len(re.findall(r'^```', text, re.MULTILINE)) % 2 == 0, path
        for target in re.findall(r'\[[^\]]+\]\(([^)]+)\)', text):
            if '://' not in target and not target.startswith('#'):
                assert (path.parent / target.split('#', 1)[0]).exists(), (path, target)
        for code in re.findall(r'```python\n(.*?)\n```', text, re.DOTALL):
            compile(code, str(path), 'exec')
    print('Shared extraction Markdown, local links and Python examples passed')


if __name__ == '__main__':
    main()
