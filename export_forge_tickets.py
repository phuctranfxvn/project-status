# /// script
# requires-python = ">=3.9"
# dependencies = ["openpyxl"]
# ///
"""Export dữ liệu tiến độ từ TMS ra xlsx (cùng định dạng file export tms.forge.ticket của TMS).

Dữ liệu căn cứ vào support ticket (tms.support.ticket):
- lọc support ticket theo project / milestone của support ticket;
- mỗi forge ticket của support ticket là một dòng;
- support ticket chưa có forge ticket vẫn có một dòng, state = "No Forge Ticket";
- milestone và ngày milestone lấy theo support ticket;
- ticket Closed (support hoặc forge) chỉ lấy khi resolution = Fixed;
- time_spent_dev tính lại = tổng giờ working hour của đúng developer của forge ticket;
  developer = developer_id, riêng ticket WIP chưa có developer_id thì lấy owner_id (Assignee).

Chạy:
    uv run export_forge_tickets.py --project kyodo-sojitz            # hỏi chọn milestone hiện hành
    uv run export_forge_tickets.py --project kyodo-sojitz --all      # cả project, không hỏi
    uv run export_forge_tickets.py --project kyodo-sojitz --all --has-milestone
    uv run export_forge_tickets.py --project kyodo-sojitz --milestone "kyodo-sojitz 1.0" -o out.xlsx
    uv run export_forge_tickets.py --domain "[('project_id.name','=','kyodo-sojitz')]"   # domain trên tms.support.ticket

Đọc thông tin kết nối từ connection.conf (url, database, http_user, http_pass, tms_user, tms_pass).
Script chỉ đọc dữ liệu, không ghi gì lên TMS.
"""
import argparse
import ast
import base64
import configparser
import datetime as dt
import os
import sys
import xmlrpc.client
from zoneinfo import ZoneInfo

from openpyxl import Workbook
from openpyxl.styles import Font

HERE = os.path.dirname(os.path.abspath(__file__))
FORGE = "tms.forge.ticket"
SUPPORT = "tms.support.ticket"
CHUNK = 500
NO_FORGE = "No Forge Ticket"  # state ghi cho support ticket chưa có forge ticket

# thứ tự cột giống file export mẫu, thêm developer và trạng thái support ở cuối
COLUMNS = [
    "id", "name", "project_id", "milestone_id", "summary", "time_spent", "time_spent_dev",
    "tms_support_ticket_id/name", "tms_support_ticket_id/workload", "development_time", "state",
    "milestone_id/date_start", "milestone_id/date_end", "wip_start", "developer_id",
    "tms_support_ticket_id/state",
]

CLOSED_FIXED_ONLY = ["|", ("state", "!=", "closed"), ("resolution", "=", "fixed")]


def load_conf(path):
    cp = configparser.ConfigParser()
    with open(path, encoding="utf-8") as f:
        cp.read_string("[c]\n" + f.read())
    return cp["c"]


class Odoo:
    def __init__(self, c):
        auth = None
        if c.get("http_user"):
            auth = "Basic " + base64.b64encode(f"{c['http_user']}:{c.get('http_pass', '')}".encode()).decode()

        class Transport(xmlrpc.client.SafeTransport):
            def send_headers(self, conn, headers):
                if auth:
                    headers.append(("Authorization", auth))
                super().send_headers(conn, headers)

        url = c["url"].rstrip("/")
        self.db, self.pwd = c["database"], c["tms_pass"]
        common = xmlrpc.client.ServerProxy(url + "/xmlrpc/2/common", transport=Transport())
        self.uid = common.authenticate(self.db, c["tms_user"], self.pwd, {})
        if not self.uid:
            sys.exit("Đăng nhập TMS thất bại: kiểm tra tms_user / tms_pass trong connection.conf")
        self.obj = xmlrpc.client.ServerProxy(url + "/xmlrpc/2/object", transport=Transport(), allow_none=True)

    def call(self, model, method, *args, **kw):
        return self.obj.execute_kw(self.db, self.uid, self.pwd, model, method, list(args), kw)

    def read(self, model, ids, fields):
        out = []
        ids = list(ids)
        for i in range(0, len(ids), CHUNK):
            out += self.call(model, "read", ids[i:i + CHUNK], fields=fields)
        return out

    def selection(self, model, field):
        return dict(self.call(model, "fields_get", [field], attributes=["selection"])[field]["selection"])


def m2o_name(v):
    return v[1] if v else None


def m2o_id(v):
    return v[0] if v else None


def to_date(s):
    return dt.datetime.strptime(s, "%Y-%m-%d") if s else None


def to_datetime(s, tz):
    # Odoo lưu datetime theo UTC; file export của TMS hiển thị theo múi giờ của user
    if not s:
        return None
    d = dt.datetime.strptime(s, "%Y-%m-%d %H:%M:%S").replace(tzinfo=dt.timezone.utc)
    return d.astimezone(tz).replace(tzinfo=None)


def choose_milestones(odoo, project):
    """Liệt kê milestone hiện hành (chưa Done) của project và hỏi người dùng chọn."""
    if not odoo.call("tms.project", "search", [("name", "=", project)], limit=1):
        sys.exit(f"Không tìm thấy project '{project}' trên TMS")
    mss = odoo.call("tms.milestone", "search_read", [("project_id.name", "=", project), ("state", "!=", "done")],
                    fields=["name", "date_start", "date_end", "state"], order="date_start, id")
    if not mss:
        sys.exit(f"Project '{project}' không có milestone hiện hành nào (tất cả đã Done). Dùng --all để lấy cả project.")
    state_label = odoo.selection("tms.milestone", "state")
    groups = odoo.call(SUPPORT, "read_group", [("milestone_id", "in", [m["id"] for m in mss])] + CLOSED_FIXED_ONLY,
                       ["milestone_id"], ["milestone_id"], lazy=False)
    counts = {g["milestone_id"][0]: g["__count"] for g in groups if g["milestone_id"]}

    print(f"Milestone hiện hành của {project}:")
    for i, m in enumerate(mss, 1):
        print(f"  {i}. {m['name']:<30} {m['date_start'] or '?'} → {m['date_end'] or '?'}  "
              f"[{state_label.get(m['state'], m['state'])}]  {counts.get(m['id'], 0)} support ticket")
    if not sys.stdin.isatty():
        sys.exit("Không chạy tương tác được: dùng --milestone hoặc --all")
    while True:
        ans = input("Chọn milestone (vd 1 hoặc 1,3 hoặc 'a' = tất cả ở trên): ").strip().lower()
        if ans in ("a", "all", "*"):
            return mss
        try:
            picks = sorted({int(x) for x in ans.replace(" ", ",").split(",") if x})
            if picks and all(1 <= k <= len(mss) for k in picks):
                return [mss[k - 1] for k in picks]
        except ValueError:
            pass
        print(f"  Nhập số từ 1 đến {len(mss)}, cách nhau bằng dấu phẩy, hoặc 'a'.")


def build_domain(args, odoo):
    """Domain trên tms.support.ticket."""
    if args.domain:
        return ast.literal_eval(args.domain) + CLOSED_FIXED_ONLY
    dom = []
    if args.project:
        dom.append(("project_id.name", "=", args.project))
    if args.milestone:
        dom.append(("milestone_id.name", "ilike", args.milestone))
    elif args.project and not args.all:
        picked = choose_milestones(odoo, args.project)
        print("Export: " + ", ".join(m["name"] for m in picked))
        dom.append(("milestone_id", "in", [m["id"] for m in picked]))
        args.picked = [m["name"] for m in picked]
    if args.has_milestone:
        dom.append(("milestone_id", "!=", False))
    if not dom:
        sys.exit("Cần --project, --milestone hoặc --domain")
    return dom + CLOSED_FIXED_ONLY


def safe(name):
    # bỏ ký tự không dùng được trong tên file, gộp khoảng trắng thừa
    return " ".join("".join("_" if ch in '\\/:*?"<>|' else ch for ch in str(name)).split())


def default_filename(args, records):
    """[Tên Project]-[Milestone]-[yyyy-mm-dd HHMMSS].xlsx, ghi ở thư mục hiện tại."""
    projects = sorted({m2o_name(r["project_id"]) for r in records} - {None})
    project = args.project or ("+".join(projects) if len(projects) <= 3 else "nhieu-project")
    names = getattr(args, "picked", None)
    if names:
        milestone = "+".join(names)
    elif args.milestone:
        milestone = args.milestone
    else:
        milestone = "all"
    return os.path.join(os.getcwd(), f"{safe(project)}-{safe(milestone)}-{dt.datetime.now():%Y-%m-%d %H%M%S}.xlsx")


def developer(t):
    """developer_id; ticket WIP chưa có developer_id thì lấy owner_id (Assignee)."""
    if t["developer_id"]:
        return t["developer_id"]
    if t["state"] == "wip":
        return t["owner_id"]
    return False


def dev_hours(t, wh_by_id):
    """time_spent_dev đúng = tổng duration_hour của working hour do chính developer ghi."""
    dev_uid = m2o_id(developer(t))
    if not dev_uid:
        return 0
    return round(sum(wh_by_id[w]["duration_hour"] for w in t["tms_working_hour_ids"]
                     if w in wh_by_id and m2o_id(wh_by_id[w]["user_id"]) == dev_uid), 2)


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--project", help="tên project TMS, ví dụ kyodo-sojitz")
    ap.add_argument("--milestone", help="lọc milestone của support ticket theo tên (ilike)")
    ap.add_argument("--all", action="store_true", help="lấy cả project, không hỏi chọn milestone")
    ap.add_argument("--has-milestone", action="store_true", help="bỏ các support ticket không có milestone")
    ap.add_argument("--domain", help="domain Odoo tuỳ ý trên tms.support.ticket, ghi đè --project/--milestone")
    ap.add_argument("--conf", default=os.path.join(HERE, "connection.conf"))
    ap.add_argument("-o", "--output", help="file xlsx đầu ra (mặc định ./[Project]-[Milestone]-[yyyy-mm-dd HHMMSS].xlsx)")
    args = ap.parse_args()

    odoo = Odoo(load_conf(args.conf))
    domain = build_domain(args, odoo)
    tz = ZoneInfo(odoo.read("res.users", [odoo.uid], ["tz"])[0]["tz"] or "UTC")

    # 1. support ticket là gốc
    sup_ids = odoo.call(SUPPORT, "search", domain, order="name")
    if not sup_ids:
        sys.exit(f"Không tìm thấy support ticket nào với domain {domain}")
    sups = odoo.read(SUPPORT, sup_ids, ["name", "project_id", "milestone_id", "summary", "workload", "state"])
    sup_state = odoo.selection(SUPPORT, "state")

    # 2. forge ticket của các support ticket đó
    forge_ids = odoo.call(FORGE, "search", [("tms_support_ticket_id", "in", sup_ids)] + CLOSED_FIXED_ONLY, order="id desc")
    forges = odoo.read(FORGE, forge_ids, [
        "name", "milestone_id", "summary", "time_spent", "time_spent_dev", "tms_support_ticket_id",
        "development_time", "state", "wip_start", "developer_id", "owner_id", "tms_working_hour_ids",
    ])
    forge_state = odoo.selection(FORGE, "state")
    forge_by_sup = {}
    for t in forges:
        forge_by_sup.setdefault(m2o_id(t["tms_support_ticket_id"]), []).append(t)

    wh_ids = [w for t in forges for w in t["tms_working_hour_ids"]]
    wh_by_id = {h["id"]: h for h in odoo.read("tms.working.hour", wh_ids, ["user_id", "duration_hour"])}
    ms_ids = {m2o_id(s["milestone_id"]) for s in sups} - {None}
    mss = {m["id"]: m for m in odoo.read("tms.milestone", ms_ids, ["date_start", "date_end"])}

    # external id giống file export (chỉ đọc id đã có, không tạo mới)
    xmlids = {}
    if forge_ids:
        for d in odoo.call("ir.model.data", "search_read", [("model", "=", FORGE), ("res_id", "in", forge_ids)],
                           fields=["module", "name", "res_id"]):
            xmlids.setdefault(d["res_id"], f"{d['module']}.{d['name']}")

    # 3. ghi file: mỗi forge ticket một dòng, support ticket không có forge ticket một dòng
    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    ws.append(COLUMNS)
    for c in ws[1]:
        c.font = Font(bold=True)

    changed, diff_ms, no_forge, n_rows = [], [], [], 0
    for s in sups:
        ms = mss.get(m2o_id(s["milestone_id"]), {})
        sup_cols = dict(project=m2o_name(s["project_id"]), milestone=m2o_name(s["milestone_id"]),
                        name=s["name"] or None, workload=s["workload"],
                        start=to_date(ms.get("date_start")), end=to_date(ms.get("date_end")),
                        state=sup_state.get(s["state"], s["state"]))
        forge = forge_by_sup.get(s["id"], [])
        if not forge:
            no_forge.append(s["name"])
            ws.append([
                f"__export__.tms_support_ticket_{s['id']}", None, sup_cols["project"], sup_cols["milestone"],
                s["summary"] or None, 0, 0, sup_cols["name"], sup_cols["workload"], 0, NO_FORGE,
                sup_cols["start"], sup_cols["end"], None, None, sup_cols["state"],
            ])
            n_rows += 1
        for t in forge:
            dh = dev_hours(t, wh_by_id)
            if abs(dh - (t["time_spent_dev"] or 0)) > 0.005:
                changed.append((t["name"], t["time_spent_dev"], dh, m2o_name(developer(t))))
            if m2o_id(t["milestone_id"]) != m2o_id(s["milestone_id"]):
                diff_ms.append(t["name"])
            ws.append([
                xmlids.get(t["id"], f"__export__.tms_forge_ticket_{t['id']}"),
                t["name"],
                sup_cols["project"],
                sup_cols["milestone"],
                t["summary"] or None,
                t["time_spent"],
                dh,
                sup_cols["name"],
                sup_cols["workload"],
                t["development_time"],
                forge_state.get(t["state"], t["state"]),
                sup_cols["start"],
                sup_cols["end"],
                to_datetime(t["wip_start"], tz),
                m2o_name(developer(t)),
                sup_cols["state"],
            ])
            n_rows += 1

    for row in ws.iter_rows(min_row=2):
        for c in row[11:13]:
            c.number_format = "yyyy-mm-dd"
        row[13].number_format = "yyyy-mm-dd hh:mm:ss"
    ws.column_dimensions["A"].width = 30.7

    out = args.output or default_filename(args, sups)
    wb.save(out)
    print(f"Đã ghi {len(sups)} support ticket, {len(forges)} forge ticket ({n_rows} dòng) vào {out}")
    if no_forge:
        print(f"{len(no_forge)} support ticket chưa có forge ticket: {', '.join(no_forge)}")
    if diff_ms:
        print(f"Lưu ý: {len(diff_ms)} forge ticket có milestone khác support ticket, đã tính theo milestone "
              f"của support: {', '.join(diff_ms)}")
    by_owner = [t["name"] for t in forges if not t["developer_id"] and developer(t)]
    if by_owner:
        print(f"{len(by_owner)} ticket WIP chưa có developer, lấy Assignee (owner_id) làm developer: {', '.join(by_owner)}")
    no_dev = [t["name"] for t in forges if not developer(t) and t["tms_working_hour_ids"]]
    if no_dev:
        print(f"Cảnh báo: {len(no_dev)} forge ticket có giờ nhưng chưa có developer, time_spent_dev = 0: {', '.join(no_dev)}")
    if changed:
        print(f"{len(changed)} forge ticket có time_spent_dev khác với giá trị trên TMS:")
        for name, old, new, dev in changed:
            print(f"  {name}: TMS {old} -> tính lại {new} (developer: {dev})")


if __name__ == "__main__":
    main()
