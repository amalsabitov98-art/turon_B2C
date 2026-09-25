import argparse
import json
from pathlib import Path


def main():
    parser = argparse.ArgumentParser(description="Build the static dashboard")
    parser.add_argument("--sales", default="data/deals.json")
    parser.add_argument("--marketing", default="data/marketing.json")
    parser.add_argument("--template", default="prototype/dashboard.tpl.html")
    parser.add_argument("--out", default="index.html")
    args = parser.parse_args()

    payload = json.loads(Path(args.sales).read_text(encoding="utf-8"))
    marketing_path = Path(args.marketing)
    payload["marketing"] = (
        json.loads(marketing_path.read_text(encoding="utf-8"))
        if marketing_path.exists()
        else {"exports": [], "issues": []}
    )

    template = Path(args.template).read_text(encoding="utf-8")
    output = template.replace(
        "/*__DATA__*/null",
        json.dumps(payload, ensure_ascii=False, separators=(",", ":")),
    )
    Path(args.out).write_text(output, encoding="utf-8")
    print("ok", len(output), "байт")


if __name__ == "__main__":
    main()
