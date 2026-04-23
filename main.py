from __future__ import annotations

import argparse
import json
import logging
import os
import re
import subprocess
import sys
import time
from dataclasses import dataclass
from pathlib import Path
from typing import Optional

import pandas as pd

DEFAULT_INPUT_SHEET = "工作表1"
DEFAULT_OUTPUT_SHEET = "工作表1"
DEFAULT_KEY_COLUMN = "KEY"
DEFAULT_DEVICE_COLUMN = "Device ID"
DEFAULT_OUTPUT_NAME = "B.xlsx"
DEFAULT_TIMEOUT_SECONDS = 20
DEFAULT_LOG_DIR = "debug_logs"


@dataclass
class ProcessResult:
    admin_password: Optional[str]
    power_user_password: Optional[str]
    raw_output: str
    error_output: str
    return_code: int
    duration_seconds: float


class PasswordGeneratorError(Exception):
    pass


def setup_logging(base_dir: Path, verbose: bool = True) -> Path:
    base_dir.mkdir(parents=True, exist_ok=True)
    log_file = base_dir / "tsl_batch_generator_debug.log"

    handlers: list[logging.Handler] = [
        logging.FileHandler(log_file, encoding="utf-8"),
    ]
    if verbose:
        handlers.append(logging.StreamHandler(sys.stdout))

    logging.basicConfig(
        level=logging.DEBUG,
        format="%(asctime)s [%(levelname)s] %(message)s",
        handlers=handlers,
        force=True,
    )
    logging.debug("Logging initialized. log_file=%s", log_file)
    return log_file



def normalize_key(value: object) -> str:
    if value is None:
        return ""
    if isinstance(value, float) and pd.isna(value):
        return ""
    return str(value).strip()



def safe_filename(text: str) -> str:
    return re.sub(r"[^A-Za-z0-9._-]+", "_", text)[:120] or "unknown"



def extract_password(output_text: str, label: str) -> Optional[str]:
    patterns = [
        rf"{re.escape(label)}\s*Password\s*[:=]\s*(\S+)",
        rf"{re.escape(label)}\s*[:=]\s*(\S+)",
        rf"{re.escape(label)}.*?Password.*?[:=]\s*(\S+)",
    ]
    for pattern in patterns:
        match = re.search(pattern, output_text, re.IGNORECASE | re.DOTALL)
        if match:
            return match.group(1).strip()
    return None



def write_text_file(path: Path, content: str) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(content, encoding="utf-8", errors="ignore")



def run_generator(
    exe_path: Path,
    key_value: str,
    timeout: int,
    debug_case_dir: Path,
) -> ProcessResult:
    if not exe_path.exists():
        raise PasswordGeneratorError(f"找不到密碼產生器: {exe_path}")

    creationflags = 0
    if sys.platform.startswith("win"):
        creationflags = getattr(subprocess, "CREATE_NO_WINDOW", 0)

    command = [str(exe_path)]
    meta = {
        "exe_path": str(exe_path),
        "cwd": str(exe_path.parent),
        "timeout": timeout,
        "key_length": len(key_value),
        "platform": sys.platform,
        "command": command,
    }
    write_text_file(debug_case_dir / "meta.json", json.dumps(meta, ensure_ascii=False, indent=2))

    logging.debug("Launching generator. exe=%s key=%s timeout=%s", exe_path, key_value, timeout)
    start_ts = time.perf_counter()

    try:
        process = subprocess.Popen(
            command,
            stdin=subprocess.PIPE,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True,
            encoding="utf-8",
            errors="ignore",
            cwd=str(exe_path.parent),
            creationflags=creationflags,
        )

        stdout_text, stderr_text = process.communicate(input=key_value + "\n", timeout=timeout)
        duration = time.perf_counter() - start_ts

    except subprocess.TimeoutExpired as exc:
        process.kill()
        stdout_text, stderr_text = process.communicate()
        duration = time.perf_counter() - start_ts
        write_text_file(debug_case_dir / "stdout_timeout.txt", stdout_text)
        write_text_file(debug_case_dir / "stderr_timeout.txt", stderr_text)
        raise PasswordGeneratorError(
            f"執行逾時，KEY={key_value}，timeout={timeout}秒"
        ) from exc
    except OSError as exc:
        duration = time.perf_counter() - start_ts
        raise PasswordGeneratorError(f"無法啟動密碼產生器: {exc}") from exc

    write_text_file(debug_case_dir / "stdout.txt", stdout_text)
    write_text_file(debug_case_dir / "stderr.txt", stderr_text)
    write_text_file(
        debug_case_dir / "summary.txt",
        "\n".join(
            [
                f"return_code={process.returncode}",
                f"duration_seconds={duration:.3f}",
                f"stdout_length={len(stdout_text)}",
                f"stderr_length={len(stderr_text)}",
            ]
        ),
    )

    admin_password = extract_password(stdout_text, "Admin")
    power_user_password = extract_password(stdout_text, "Power User")

    logging.debug(
        "Generator finished. return_code=%s duration=%.3f admin_found=%s power_user_found=%s",
        process.returncode,
        duration,
        bool(admin_password),
        bool(power_user_password),
    )

    return ProcessResult(
        admin_password=admin_password,
        power_user_password=power_user_password,
        raw_output=stdout_text,
        error_output=stderr_text,
        return_code=process.returncode,
        duration_seconds=duration,
    )



def validate_required_columns(df: pd.DataFrame, device_column: str, key_column: str) -> None:
    missing = [col for col in [device_column, key_column] if col not in df.columns]
    if missing:
        raise ValueError(f"Excel 缺少必要欄位: {missing}")



def process_excel(
    input_excel: Path,
    exe_path: Path,
    output_excel: Path,
    debug_root: Path,
    input_sheet: str = DEFAULT_INPUT_SHEET,
    output_sheet: str = DEFAULT_OUTPUT_SHEET,
    device_column: str = DEFAULT_DEVICE_COLUMN,
    key_column: str = DEFAULT_KEY_COLUMN,
    timeout: int = DEFAULT_TIMEOUT_SECONDS,
) -> Path:
    logging.info("開始讀取 Excel: %s", input_excel)
    df = pd.read_excel(input_excel, sheet_name=input_sheet)
    validate_required_columns(df, device_column, key_column)
    logging.debug("Excel loaded. rows=%s columns=%s", len(df), list(df.columns))

    results: list[dict[str, object]] = []

    for idx, row in df.iterrows():
        device_id = str(row.get(device_column, "")).strip()
        key_value = normalize_key(row.get(key_column, ""))
        case_name = f"{idx+1:04d}_{safe_filename(device_id or 'NO_DEVICE')}_{safe_filename(key_value)}"
        case_dir = debug_root / case_name
        case_dir.mkdir(parents=True, exist_ok=True)

        logging.info("處理第 %s 筆, Device ID=%s, KEY=%s", idx + 1, device_id, key_value)
        write_text_file(
            case_dir / "input.json",
            json.dumps(
                {
                    "row_number": idx + 2,
                    "device_id": device_id,
                    "key": key_value,
                    "key_length": len(key_value),
                },
                ensure_ascii=False,
                indent=2,
            ),
        )

        result_row = {
            device_column: device_id,
            key_column: key_value,
            "Admin Password": "",
            "Power User Password": "",
            "Status": "",
            "Message": "",
            "Return Code": "",
            "Duration (s)": "",
            "Debug Folder": str(case_dir),
        }

        if not key_value:
            result_row["Status"] = "FAIL"
            result_row["Message"] = "KEY為空"
            results.append(result_row)
            logging.warning("第 %s 筆失敗: KEY為空", idx + 1)
            continue

        if len(key_value) != 20:
            result_row["Status"] = "FAIL"
            result_row["Message"] = f"KEY長度不符，實際長度={len(key_value)}"
            results.append(result_row)
            logging.warning("第 %s 筆失敗: KEY長度不符", idx + 1)
            continue

        try:
            process_result = run_generator(
                exe_path=exe_path,
                key_value=key_value,
                timeout=timeout,
                debug_case_dir=case_dir,
            )
            result_row["Power User Password"] = process_result.power_user_password or ""
            result_row["Admin Password"] = process_result.admin_password or ""
            result_row["Return Code"] = process_result.return_code
            result_row["Duration (s)"] = round(process_result.duration_seconds, 3)

            parsed = {
                "admin_password": process_result.admin_password,
                "power_user_password": process_result.power_user_password,
                "return_code": process_result.return_code,
                "duration_seconds": process_result.duration_seconds,
            }
            write_text_file(case_dir / "parsed.json", json.dumps(parsed, ensure_ascii=False, indent=2))

            if process_result.admin_password:
                result_row["Status"] = "OK"
                result_row["Message"] = ""
                logging.info("第 %s 筆成功", idx + 1)
            else:
                result_row["Status"] = "FAIL"
                err_msg = "抓不到 Admin Password"
                if process_result.error_output.strip():
                    err_msg += f" | STDERR: {process_result.error_output.strip()}"
                result_row["Message"] = err_msg
                logging.error("第 %s 筆失敗: %s", idx + 1, err_msg)
        except Exception as exc:
            result_row["Status"] = "FAIL"
            result_row["Message"] = str(exc)
            logging.exception("第 %s 筆執行異常", idx + 1)
            write_text_file(case_dir / "exception.txt", f"{type(exc).__name__}: {exc}")

        results.append(result_row)

    output_df = pd.DataFrame(results)
    output_excel.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(output_excel, engine="openpyxl") as writer:
        output_df.to_excel(writer, sheet_name=output_sheet, index=False)

    logging.info("處理完成，輸出檔案: %s", output_excel)
    logging.info("Debug logs 位置: %s", debug_root)
    return output_excel



def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="TSL Password Batch Generator Debug Version")
    parser.add_argument("--input", default="A.xlsx", help="輸入 Excel 檔案路徑")
    parser.add_argument("--exe", default="TSL_password_generator.exe", help="密碼產生器 exe 路徑")
    parser.add_argument("--output", default="B.xlsx", help="輸出 Excel 檔案路徑")
    parser.add_argument("--input-sheet", default=DEFAULT_INPUT_SHEET, help="輸入工作表名稱")
    parser.add_argument("--output-sheet", default=DEFAULT_OUTPUT_SHEET, help="輸出工作表名稱")
    parser.add_argument("--device-column", default=DEFAULT_DEVICE_COLUMN, help="Device ID 欄位名稱")
    parser.add_argument("--key-column", default=DEFAULT_KEY_COLUMN, help="KEY 欄位名稱")
    parser.add_argument("--timeout", default=DEFAULT_TIMEOUT_SECONDS, type=int, help="每筆執行逾時秒數")
    parser.add_argument("--debug-dir", default=DEFAULT_LOG_DIR, help="debug log 輸出資料夾")
    return parser.parse_args()



def main() -> int:
    args = parse_args()
    input_excel = Path(args.input).resolve()
    exe_path = Path(args.exe).resolve()
    output_excel = Path(args.output).resolve()
    debug_root = output_excel.parent / args.debug_dir

    setup_logging(debug_root, verbose=True)

    logging.debug("Arguments: %s", vars(args))
    logging.debug("Current working directory: %s", os.getcwd())
    logging.debug("Resolved input_excel=%s", input_excel)
    logging.debug("Resolved exe_path=%s", exe_path)
    logging.debug("Resolved output_excel=%s", output_excel)
    logging.debug("Resolved debug_root=%s", debug_root)

    try:
        process_excel(
            input_excel=input_excel,
            exe_path=exe_path,
            output_excel=output_excel,
            debug_root=debug_root,
            input_sheet=args.input_sheet,
            output_sheet=args.output_sheet,
            device_column=args.device_column,
            key_column=args.key_column,
            timeout=args.timeout,
        )
        print(f"完成，請查看輸出檔: {output_excel}")
        print(f"Debug log 資料夾: {debug_root}")
        return 0
    except Exception as exc:
        logging.exception("主程式執行失敗")
        print(f"ERROR: {exc}")
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
