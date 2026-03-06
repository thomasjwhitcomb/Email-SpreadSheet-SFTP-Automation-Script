#!/usr/bin/env python3
r"""
sftp_upload_password.py (STRICT + KEYRING VERSION)

This version:
- Uses PASSWORD authentication for SFTP
- Does NOT attempt to create remote directories
- Verifies the remote directory exists before upload
- Retrieves password from OS keychain via `keyring` if --password is not provided

Keyring notes:
- Windows: Credential Manager
- macOS: Keychain
- The password is stored per user account. The script must run under the same OS account
  that stored the secret.

One-time store (example):
  python -c "import keyring; keyring.set_password('rqi-sftp','116286', input('Password: '))"

Run (no password on CLI):
  python sftp_upload_password.py ^
    --host rqi1stop-sftp-preprod.rqi1stop.com ^
    --port 6239 ^
    --username 116286 ^
    --local-dir "C:/temp" ^
    --filename "required_filename.csv" ^
    --remote-dir "/incoming" ^
    --keyring-service "rqi-sftp" ^
    --verify-size

For production:
- Do NOT use --auto-add-hostkey unless you fully understand the risk.
- Prefer validating host keys via known_hosts.
"""

import os
import argparse
import logging
import hashlib
import errno
import paramiko
import keyring


logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")


def sha256_file(path: str, chunk_size: int = 65536) -> str:
    h = hashlib.sha256()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(chunk_size), b""):
            h.update(chunk)
    return h.hexdigest()


def verify_remote_dir_exists(sftp: paramiko.SFTPClient, remote_dir: str) -> None:
    """
    Verify that the remote directory exists and is accessible.
    Do NOT attempt to create it.
    Raise RuntimeError if it does not exist or is not accessible.
    """
    if not remote_dir or remote_dir in (".", "/"):
        return

    remote_dir = remote_dir.replace("\\", "/").rstrip("/")

    try:
        sftp.stat(remote_dir)
    except IOError as e:
        err = getattr(e, "errno", None)

        # Some SFTP servers return text only; we check both errno and message.
        if err == errno.ENOENT or "No such file" in str(e):
            raise RuntimeError(f"Remote directory '{remote_dir}' does not exist. Upload aborted.") from e

        if err == errno.EACCES or "Permission denied" in str(e):
            raise RuntimeError(f"No permission to access remote directory '{remote_dir}'. Upload aborted.") from e

        raise


def resolve_password(username: str, cli_password: str | None, keyring_service: str) -> str:
    """
    Resolve password using:
      1) --password if provided
      2) OS keychain via keyring (service + username)
    """
    if cli_password:
        return cli_password

    pw = keyring.get_password(keyring_service, username)
    if not pw:
        raise RuntimeError(
            "Password not provided and not found in keyring.\n"
            f"Store it with:\n"
            f'  python -c "import keyring; keyring.set_password(\'{keyring_service}\',\'{username}\', input(\'Password: \'))"'
        )
    return pw


def upload_with_password(host: str,
                         port: int,
                         username: str,
                         password: str,
                         local_dir: str,
                         filename: str,
                         remote_dir: str,
                         verify_sha256: bool = False,
                         verify_size: bool = True,
                         auto_add_host_key: bool = False,
                         timeout: int = 10) -> bool:
    if not filename:
        logging.error("Filename is required.")
        return False

    local_path = os.path.join(local_dir, filename)
    if not os.path.isfile(local_path):
        logging.error("Local file does not exist: %s", local_path)
        return False

    # Normalize remote paths (SFTP uses POSIX-style forward slashes)
    remote_dir = (remote_dir or "").replace("\\", "/").rstrip("/")
    if remote_dir in ("", "."):
        remote_dir = "/"

    remote_path = (remote_dir.rstrip("/") + "/" + filename) if remote_dir != "/" else ("/" + filename)

    client = paramiko.SSHClient()

    # Load known_hosts if present
    try:
        client.load_system_host_keys()
    except Exception:
        pass

    if auto_add_host_key:
        logging.warning("Auto-adding unknown host keys (MITM risk). For production, use known_hosts.")
        client.set_missing_host_key_policy(paramiko.AutoAddPolicy())
    else:
        client.set_missing_host_key_policy(paramiko.RejectPolicy())

    sftp = None
    try:
        logging.info("Connecting to %s:%d as %s", host, port, username)

        client.connect(
            hostname=host,
            port=port,
            username=username,
            password=password,
            timeout=timeout,
            banner_timeout=timeout,
            look_for_keys=False,
            allow_agent=False,
        )

        logging.info("Authentication (password) successful!")

        sftp = client.open_sftp()
        logging.info("Opened sftp connection")

        logging.info("Verifying remote directory: %s", remote_dir)
        verify_remote_dir_exists(sftp, remote_dir)

        logging.info("Uploading %s -> %s", local_path, remote_path)
        sftp.put(local_path, remote_path)

        if verify_size:
            local_size = os.path.getsize(local_path)
            remote_size = sftp.stat(remote_path).st_size
            logging.info("Local size: %d bytes, Remote size: %d bytes", local_size, remote_size)
            if local_size != remote_size:
                logging.error("Size mismatch after upload.")
                return False

        if verify_sha256:
            logging.info("Computing local SHA256...")
            local_hash = sha256_file(local_path)

            logging.info("Computing remote SHA256 by streaming remote file...")
            h = hashlib.sha256()
            with sftp.open(remote_path, "rb") as rf:
                for chunk in iter(lambda: rf.read(65536), b""):
                    h.update(chunk)

            remote_hash = h.hexdigest()

            logging.info("Local SHA256:  %s", local_hash)
            logging.info("Remote SHA256: %s", remote_hash)

            if local_hash != remote_hash:
                logging.error("SHA256 mismatch after upload.")
                return False

        logging.info("Upload completed and verified.")
        return True

    except paramiko.AuthenticationException:
        logging.exception("Authentication failed (password rejected).")
        return False
    except RuntimeError as e:
        logging.error(str(e))
        return False
    except paramiko.SSHException as e:
        logging.exception("SSH error: %s", e)
        return False
    finally:
        try:
            if sftp is not None:
                sftp.close()
                logging.info("sftp session closed.")
        except Exception:
            pass
        client.close()


def parse_args():
    p = argparse.ArgumentParser(description="SFTP upload using password auth (STRICT remote dir + KEYRING).")

    p.add_argument("--host", required=True)
    p.add_argument("--port", type=int, default=22)
    p.add_argument("--username", required=True)

    # Password is OPTIONAL now (prefer keyring)
    p.add_argument("--password", default=None, help="SFTP password (avoid on CLI; prefer keyring).")
    p.add_argument("--keyring-service", default="rqi-sftp",
                   help="Keyring service name used to look up password (default: rqi-sftp).")

    p.add_argument("--local-dir", required=True)
    p.add_argument("--filename", required=True, help="Exact filename required by receiver.")
    p.add_argument("--remote-dir", required=True)

    p.add_argument("--verify-sha256", action="store_true")
    p.add_argument("--no-verify-size", dest="verify_size", action="store_false")
    p.add_argument("--auto-add-hostkey", action="store_true")
    return p.parse_args()


if __name__ == "__main__":
    args = parse_args()

    try:
        password = resolve_password(args.username, args.password, args.keyring_service)
    except RuntimeError as e:
        logging.error(str(e))
        raise SystemExit(1)

    ok = upload_with_password(
        host=args.host,
        port=args.port,
        username=args.username,
        password=password,
        local_dir=args.local_dir,
        filename=args.filename,
        remote_dir=args.remote_dir,
        verify_sha256=args.verify_sha256,
        verify_size=args.verify_size,
        auto_add_host_key=args.auto_add_hostkey,
    )

    if not ok:
        raise SystemExit(1)