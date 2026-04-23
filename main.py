from __future__ import annotations

import argparse
import logging
import re
import subprocess
import sys
from dataclasses import dataclass
from pathlib import Path
from typing import Optional

import pandas as pd

DEFAULT_INPUT_SHEET = "工作表1"
DEFAULT_OUTPUT_SHEET = "工作表1"
DEFAULT_KEY_COLUMN = "KEY"
DEFAULT_DEVICE_COLUMN = "Device ID"
DEFAULT_OUTPUT_NAME = "B.xlsx"
DEFAULT_TIMEOUT_SECONDS = 15


@dataclass
class ProcessResult:
    admin_password: Optional[str]
    power_user_password: Optional[str]
    raw_output: str
    error_output: str
    return_code: int


class PasswordGeneratorError(Exception):
    pass


def setup_logging(output_dir: Path) -> None:
    output_dir.mkdir(parents=True, exist_ok=True)
    log_file = output_dir / "tsl_batch_generator.log"
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s [%(levelname)s] %(message)s",
        handlers=[
            logging.FileHandler(log_file, encoding="utf-8"),
            logging.StreamHandler(sys.stdout),
        ],
    )


def normalize_key(value: object) -> str:
    if value is None:
        return ""
    if isinstance(value, float) and pd.isna(value):
        return ""
    return str(value).strip()


def extract_password(output_text: str, label: str) -> Optional[str]:
    patterns = [
        rf"{re.escape(label)}\s*Password\s*[:=]\s*(\S+)",
        rf"{re.escape(label)}\s*[:=]\s*(\S+)",
        rf"{re.escape(label)}.*?Password.*?[:=]\s*(\S+)",
    ]
    for pattern in patterns:
        match = re.search(pattern, output_text, re.IGNORECASE)
        if match:
            return match.group(1).strip()
    return None


def run_generator(exe_path: Path, key_value: str, timeout: int) -> ProcessResult:
    if not exe_path.exists():
        raise PasswordGeneratorError(f"找不到密碼產生器: {exe_path}")

    creationflags = 0
    if sys.platform.startswith("win"):
        creationflags = subprocess.CREATE_NO_WINDOW  # type: ignore[attr-defined]

    try:
        process = subprocess.Popen(
            [str(exe_path)],
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
    except subprocess.TimeoutExpired as exc:
        process.kill()
        stdout_text, stderr_text = process.communicate()
        raise PasswordGeneratorError(
            f"執行逾時，KEY={key_value}，timeout={timeout}秒\nSTDOUT:\n{stdout_text}\nSTDERR:\n{stderr_text}"
        ) from exc
    except OSError as exc:
        raise PasswordGeneratorError(f"無法啟動密碼產生器: {exc}") from exc

    admin_password = extract_password(stdout_text, "Admin")
    power_user_password = extract_password(stdout_text, "Power User")

    return ProcessResult(
        admin_password=admin_password,
        power_user_password=power_user_password,
        raw_output=stdout_text,
        error_output=stderr_text,
        return_code=process.returncode,
    )


def validate_required_columns(df: pd.DataFrame, device_column: str, key_column: str) -> None:
    missing = [col for col in [device_column, key_column] if col not in df.columns]
    if missing:
        raise ValueError(f"Excel 缺少必要欄位: {missing}")


def process_excel(
    input_excel: Path,
    exe_path: Path,
    output_excel: Path,
    input_sheet: str = DEFAULT_INPUT_SHEET,
    output_sheet: str = DEFAULT_OUTPUT_SHEET,
    device_column: str = DEFAULT_DEVICE_COLUMN,
    key_column: str = DEFAULT_KEY_COLUMN,
    timeout: int = DEFAULT_TIMEOUT_SECONDS,
) -> Path:
    logging.info("開始讀取 Excel: %s", input_excel)
    df = pd.read_excel(input_excel, sheet_name=input_sheet)
    validate_required_columns(df, device_column, key_column)

    results: list[dict[str, object]] = []

    for idx, row in df.iterrows():
        device_id = row.get(device_column, "")
        key_value = normalize_key(row.get(key_column, ""))
        logging.info("處理第 %s 筆, Device ID=%s, KEY=%s", idx + 1, device_id, key_value)

        result_row = {
            device_column: device_id,
            key_column: key_value,
            "Admin Password": "",
            "Power User Password": "",
            "Status": "",
            "Message": "",
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
            process_result = run_generator(exe_path=exe_path, key_value=key_value, timeout=timeout)
            result_row["Power User Password"] = process_result.power_user_password or ""
            result_row["Admin Password"] = process_result.admin_password or ""

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
        except Exception as exc:  # noqa: BLE001
            result_row["Status"] = "FAIL"
            result_row["Message"] = str(exc)
            logging.exception("第 %s 筆執行異常", idx + 1)

        results.append(result_row)

    output_df = pd.DataFrame(results)
    output_excel.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(output_excel, engine="openpyxl") as writer:
        output_df.to_excel(writer, sheet_name=output_sheet, index=False)

    logging.info("處理完成，輸出檔案: %s", output_excel)
    return output_excel


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="TSL Password Batch Generator")
    parser.add_argument("--input", default="A.xlsx", help="輸入 Excel 檔案路徑")
    parser.add_argument("--exe", default="TSL_password_generator.exe", help="密碼產生器 exe 路徑")
    parser.add_argument("--output", default=DEFAULT_OUTPUT_NAME, help="輸出 Excel 檔案路徑")
    parser.add_argument("--input-sheet", default=DEFAULT_INPUT_SHEET, help="輸入工作表名稱")
    parser.add_argument("--output-sheet", default=DEFAULT_OUTPUT_SHEET, help="輸出工作表名稱")
    parser.add_argument("--device-column", default=DEFAULT_DEVICE_COLUMN, help="Device ID 欄位名稱")
    parser.add_argument("--key-column", default=DEFAULT_KEY_COLUMN, help="KEY 欄位名稱")
    parser.add_argument("--timeout", default=DEFAULT_TIMEOUT_SECONDS, type=int, help="每筆執行逾時秒數")
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    input_excel = Path(args.input).resolve()
    exe_path = Path(args.exe).resolve()
    output_excel = Path(args.output).resolve()
    setup_logging(output_excel.parent)

    try:
        process_excel(
            input_excel=input_excel,
            exe_path=exe_path,
            output_excel=output_excel,
            input_sheet=args.input_sheet,
            output_sheet=args.output_sheet,
            device_column=args.device_column,
            key_column=args.key_column,
            timeout=args.timeout,
        )
        return 0
    except Exception as exc:  # noqa: BLE001
        logging.exception("主程式執行失敗")
        print(f"ERROR: {exc}")
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
