"""One-shot patch: wire self_heal.py into both bat files."""
import re
from pathlib import Path

BASE = Path(__file__).parent
BATS = ["update_coaching_dashboard_auto.bat", "update_coaching_dashboard.bat"]

for bat_name in BATS:
    bat = BASE / bat_name
    c = bat.read_text(encoding="utf-8")

    # 1. Replace all pull/build step invocations:
    #    py -3 script.py 2>"%MASTERLIST_DIR%\step_err.tmp"
    # → py -3 "%MASTERLIST_DIR%\self_heal.py" run-step script.py
    c = re.sub(
        r'py -3 (\S+\.py) 2>"%MASTERLIST_DIR%\\step_err\.tmp"',
        r'py -3 "%MASTERLIST_DIR%\\self_heal.py" run-step \1',
        c
    )

    # 2. Remove standalone fix_footgun line (now handled inside self_heal.py)
    c = re.sub(r'py -3 "%MASTERLIST_DIR%\\fix_footgun\.py"\r?\n', '', c)

    # 3. Add heal-drops step before "Syncing latest dashboard repo changes"
    heal_block = (
        "echo.\r\n"
        "echo ========================================\r\n"
        "echo Self-heal: checking for count drops\r\n"
        "echo ========================================\r\n"
        'py -3 "%MASTERLIST_DIR%\\self_heal.py" heal-drops\r\n'
        "\r\n"
    )
    c = c.replace(
        "echo Syncing latest dashboard repo changes...",
        heal_block + "echo Syncing latest dashboard repo changes..."
    )

    # 4. Add git push retry (replace bare "git push\r\nif errorlevel 1 goto :fail")
    c = c.replace(
        "git push\r\nif errorlevel 1 goto :fail\r\n\r\npy -3",
        (
            "git push\r\n"
            "if errorlevel 1 (\r\n"
            "    echo [self-heal] Push failed, retrying in 30 seconds...\r\n"
            "    timeout /t 30 /nobreak >nul\r\n"
            "    git push\r\n"
            "    if errorlevel 1 (\r\n"
            '        py -3 "%MASTERLIST_DIR%\\log_step.py" step "Git Push" "git push" 1\r\n'
            "        goto :fail\r\n"
            "    )\r\n"
            ")\r\n"
            "\r\npy -3"
        )
    )

    bat.write_bytes(c.encode("utf-8"))
    print(f"Patched {bat_name}")

print("Done.")
