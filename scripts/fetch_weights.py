"""从 ModelScope 下载模型（宿主机访问不了 huggingface.co）。只用标准库；支持断点续传和 sha256 校验。
    python scripts/fetch_weights.py Qwen/Qwen2.5-Coder-1.5B-Instruct models
"""
import hashlib, json, os, subprocess, sys, time, urllib.request

API = "https://www.modelscope.cn/api/v1/models/{repo}/repo/files?Revision=master&Recursive=true"
RAW = "https://www.modelscope.cn/models/{repo}/resolve/master/{path}"


def sha256(p):
    h = hashlib.sha256()
    with open(p, "rb") as f:
        for b in iter(lambda: f.read(1 << 22), b""):
            h.update(b)
    return h.hexdigest()


def fetch(repo, root):
    dst = os.path.join(root, repo.split("/")[-1])
    os.makedirs(dst, exist_ok=True)
    files = json.load(urllib.request.urlopen(API.format(repo=repo), timeout=60))["Data"]["Files"]
    for f in files:
        if f.get("Type") == "tree":
            continue
        p = os.path.join(dst, f["Path"])
        os.makedirs(os.path.dirname(p), exist_ok=True)
        size = f["Size"]
        for attempt in range(20):
            have = os.path.getsize(p) if os.path.exists(p) else 0
            if have == size:
                break
            if have > size:
                os.remove(p)
            # urllib 在 ModelScope 大文件上会卡住（0 字节不动），改用系统 curl：自动跟随跳转 + 断点续传
            rc = subprocess.call(["curl", "-sS", "-L", "-C", "-", "--retry", "5", "--speed-limit", "102400", "--speed-time", "60",
                                  "-o", p, RAW.format(repo=repo, path=f["Path"])])
            print(f"  curl rc={rc} {f['Path']} {os.path.getsize(p) / 2**30:.2f}/{size / 2**30:.2f} GiB", flush=True)
        if f.get("Sha256") and sha256(p) != f["Sha256"]:
            print("SHA MISMATCH", p, flush=True); os.remove(p); sys.exit(2)
        print("ok", f["Path"], size, flush=True)
    print("DONE", repo, dst, flush=True)


if __name__ == "__main__":
    for repo in sys.argv[1:-1]:
        fetch(repo, sys.argv[-1])
