#!/usr/bin/env python3
"""Call the confined remote registration via the manually opened eksprod route."""
import argparse
import base64
import hashlib
import json
import os
from pathlib import Path
import re
import subprocess
import tempfile

ROOT = Path(__file__).resolve().parents[1]
PREPARED = ROOT / "deploy/prepared"


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("mode", choices=("preflight", "apply", "export-secret"))
    args = parser.parse_args()
    access = subprocess.run(["/root/eksprod", "status"], check=True, capture_output=True, text=True)
    status = json.loads(access.stdout)
    if status.get("eksprod_sudo") != "yes" or status.get("origin") != "manual":
        raise SystemExit("Manual eksprod administrative access is not open")
    policy = (PREPARED / "client-registration.json").read_bytes()
    source = (ROOT / "deploy/register_idp_client.py").read_bytes()
    code = "import base64;exec(compile(base64.b64decode(%r),'smeta-register','exec'))" % base64.b64encode(source).decode()
    argv = ["/root/eksprod", "sudo", "--runtime-seconds", "180", "--", "/usr/bin/python3", "-I", "-S", "-B",
            "-c", code, args.mode, base64.b64encode(policy).decode(), status["activation_id"]]
    if args.mode != "export-secret":
        result = subprocess.run(argv, capture_output=True, text=True)
        if result.returncode:
            print(result.stderr.strip())
            raise SystemExit(result.returncode)
        receipt = json.loads(result.stdout)
        if receipt.get("client_id") != "cii-smeta-parser":
            raise SystemExit("Unexpected registration result")
        receipt["program_sha256"] = hashlib.sha256(source).hexdigest()
        (PREPARED / ("idp-" + args.mode + ".json")).write_text(json.dumps(receipt, ensure_ascii=False, indent=2) + "\n")
        print(json.dumps(receipt, ensure_ascii=False))
        return
    proof = json.loads((PREPARED / "idp-apply.json").read_text())
    if proof.get("registered") is not True or proof.get("other_clients_unchanged") is not True:
        raise SystemExit("Verified own client registration is required")
    folder = Path("/etc/smeta-parser")
    descriptor, temporary = tempfile.mkstemp(prefix=".client-secret-", dir=folder)
    try:
        with os.fdopen(descriptor, "wb") as output:
            os.fchmod(output.fileno(), 0o600)
            result = subprocess.run(argv, stdout=output, stderr=subprocess.PIPE)
        if result.returncode:
            raise SystemExit("Own credential export failed; no credential installed")
        value = Path(temporary).read_bytes()
        if not re.fullmatch(rb"[A-Za-z0-9_-]{24,256}", value):
            raise SystemExit("Own credential response rejected")
        target = folder / "client-secret"
        if target.exists():
            if target.is_symlink() or target.read_bytes() != value:
                raise SystemExit("Existing credential differs; refusing replacement")
        else:
            os.replace(temporary, target)
        marker = folder / "client-registration.ok"
        public_proof = {k: proof[k] for k in ("schema", "client_id", "client_uuid", "registered", "issuer",
                                             "display_name", "policy_sha256", "other_clients_unchanged")}
        fd = os.open(marker, os.O_WRONLY | os.O_CREAT | os.O_EXCL, 0o600)
        with os.fdopen(fd, "w") as stream:
            json.dump(public_proof, stream, ensure_ascii=False)
            stream.write("\n")
        print("Own client credential installed in a 0600 file; registration proof recorded")
    finally:
        Path(temporary).unlink(missing_ok=True)


if __name__ == "__main__":
    main()
