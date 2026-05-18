from __future__ import annotations

import subprocess
import sys
from datetime import datetime
from pathlib import Path

REPO_DIR = Path(__file__).resolve().parent
LOG_DIR = REPO_DIR / 'logs'
LOG_FILE = LOG_DIR / 'task.log'
CONVERT_SCRIPT = REPO_DIR / 'convert.py'
BRANCH = 'main'
REMOTE = 'origin'
COMMIT_PREFIX = 'actualizacion automatica'
FILES_TO_COMMIT = [
    REPO_DIR / 'data' / 'news.json',
]


def log(message: str) -> None:
    LOG_DIR.mkdir(parents=True, exist_ok=True)
    timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    line = f'[{timestamp}] {message}'
    print(line)
    with LOG_FILE.open('a', encoding='utf-8') as f:
        f.write(line + '\n')


def run(cmd: list[str], check: bool = True) -> subprocess.CompletedProcess[str]:
    log('Ejecutando: ' + ' '.join(cmd))
    proc = subprocess.run(
        cmd,
        cwd=REPO_DIR,
        capture_output=True,
        text=True,
        encoding='utf-8',
        errors='replace',
    )
    if proc.stdout.strip():
        with LOG_FILE.open('a', encoding='utf-8') as f:
            f.write(proc.stdout)
            if not proc.stdout.endswith('\n'):
                f.write('\n')
    if proc.stderr.strip():
        with LOG_FILE.open('a', encoding='utf-8') as f:
            f.write(proc.stderr)
            if not proc.stderr.endswith('\n'):
                f.write('\n')
    if check and proc.returncode != 0:
        raise RuntimeError(
            f"Falló el comando ({proc.returncode}): {' '.join(cmd)}"
        )
    return proc


def ensure_git_repo() -> None:
    git_dir = REPO_DIR / '.git'
    if not git_dir.exists():
        raise FileNotFoundError(
            f'No se encontró .git en {REPO_DIR}. Copia este archivo dentro del repo correcto.'
        )


def run_convert_script() -> None:
    if not CONVERT_SCRIPT.exists():
        raise FileNotFoundError(
            f'No se encontró {CONVERT_SCRIPT.name} en {REPO_DIR}. '
            'Asegúrate de poner este archivo en la misma carpeta que convert.py.'
        )
    run([sys.executable, str(CONVERT_SCRIPT)])


def sync_branch() -> None:
    run(['git', 'pull', '--ff-only', REMOTE, BRANCH])


def stage_generated_files() -> None:
    paths = [str(path.relative_to(REPO_DIR)) for path in FILES_TO_COMMIT if path.exists()]
    if not paths:
        raise FileNotFoundError('No se encontraron archivos generados para agregar a Git.')
    run(['git', 'add', *paths])


def has_staged_changes() -> bool:
    proc = run(['git', 'diff', '--cached', '--quiet'], check=False)
    return proc.returncode == 1


def commit_and_push() -> None:
    message = f'{COMMIT_PREFIX}: {datetime.now().strftime("%Y-%m-%d %H:%M")}'
    run(['git', 'commit', '-m', message])
    run(['git', 'push', REMOTE, BRANCH])


def main() -> int:
    try:
        log('================ INICIO ================')
        ensure_git_repo()
        sync_branch()
        run_convert_script()
        stage_generated_files()

        if has_staged_changes():
            log('Se detectaron cambios. Haciendo commit y push...')
            commit_and_push()
        else:
            log('No hubo cambios. No se hará commit ni push.')

        log('================= FIN =================')
        return 0
    except Exception as exc:
        log(f'ERROR: {exc}')
        log('============= FIN CON ERROR ============')
        return 1


if __name__ == '__main__':
    raise SystemExit(main())
