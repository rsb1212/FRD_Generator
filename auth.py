import os


def authenticate_user(username: str, password: str) -> bool:
    expected_username = os.getenv("FRD_ADMIN_USERNAME", "admin").strip()
    expected_password = os.getenv("FRD_ADMIN_PASSWORD", "admin123").strip()
    return username.strip() == expected_username and password.strip() == expected_password
