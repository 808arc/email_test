import copy
import random
import time
from pathlib import Path
import pandas as pd
from verify_email import verify_email


class FileNotFound(Exception):
    pass


class RecordEmail:
    def __init__(self, file_paths):
        self.file_paths = file_paths
        self.data = self.read_excel_files()
        self.validate_emails()
        self.save_updated_files()
        self.save_updated_files_to_excel()

    @classmethod
    def dict_init(cls, custom_dict: dict[str, str] = None) -> dict[str, str]:
        return copy.deepcopy(custom_dict if custom_dict is not None else {})

    def read_excel_files(self):
        all_data = self.dict_init()
        for i, path in enumerate(self.file_paths):
            if not path.exists():
                raise FileNotFoundError(f"{path} not found")
            df = pd.read_excel(path)

            df = df.drop_duplicates(keep="first")

            key_col = df.columns[0]
            for email in df[key_col]:
                if email not in all_data:
                    all_data[email] = {"file": i, "status": "Not validated"}
                elif all_data[email]["file"] < i:
                    all_data[email][
                        "status"
                    ] += f", Was in file {all_data[email]['file'] + 1}"

        return all_data

    def validate_emails(self):
        total_emails = len(self.data)
        emails_checked = 0  # Initialize the counter for checked emails
        for email in self.data:
            try:
                is_valid = verify_email(str(email))
                validity = "Valid" if is_valid else "Invalid"
                if self.data[email]["status"] == "Not validated":
                    self.data[email]["status"] = validity
                else:
                    self.data[email]["status"] = (
                        f"{validity}, " + self.data[email]["status"]
                    )
            except Exception as e:
                error_msg = f"Error: {str(e)}"
                if self.data[email]["status"] == "Not validated":
                    self.data[email]["status"] = error_msg
                else:
                    self.data[email]["status"] = (
                        error_msg + ", " + self.data[email]["status"]
                    )
            emails_checked += 1  # Increment the counter for checked emails
            print(
                f"Total emails checked: {emails_checked}/{total_emails}", end="\r"
            )  # Print the progress in the same line

            if emails_checked % 10 == 0:  # Add random await every 10 emails
                await_time = random.uniform(1, 3)  # Random await time between 1 - 3 seconds
                time.sleep(await_time)

    def save_updated_files(self):
        for path in self.file_paths:
            df = pd.read_excel(path)
            key_col = df.columns[0]
            df["Status"] = df[key_col].map(
                lambda email: (
                    self.data[email]["status"] if email in self.data else "Unknown"
                )
            )
            updated_path = path.with_stem(f"{path.stem}_updated")
            df.to_excel(updated_path, index=False)
            print(f"Updated file saved: {updated_path}")

    def save_updated_files_to_excel(self):
        all_emails = list(self.data.keys())
        all_statuses = [self.data[email]["status"] if email in self.data else "Unknown" for email in all_emails]
        df = pd.DataFrame({"Email": all_emails, "Status": all_statuses})
        updated_path = Path("validated_emails.xlsx")  # Specify the desired file name
        df.to_excel(updated_path, index=False)
        print(f"Updated file saved: {updated_path}")


def main():
    current_dir = Path(__file__).parent
    file_paths = list(current_dir.glob("tabs/*.xlsx"))
    record = RecordEmail(file_paths)

    print(f"\nTotal rows processed: {len(record.data)}")
    print("Sample of processed data:")
    for key, value in list(record.data.items())[:5]:
        print(f"{key}: {value}")


if __name__ == "__main__":
    main()
