from fastapi.testclient import TestClient
import ironcalc as ic, os, sys, tempfile
sys.path.insert(0, "."); from server.app import app   # リポジトリのルートで python server/test_app.py
c = TestClient(app)
r = c.post("/api/open", files={"file": open("sample_workpaper.xlsx","rb")}); sid = r.headers["X-Session-Id"]
cli = ic.UserModel.from_bytes(bytes(r.content), "en")
cli.set_user_input(1, 10, 2, "1"); diffs = bytes(cli.flush_send_queue())
out = os.path.join(tempfile.gettempdir(), "out.xlsx")
r = c.post(f"/api/sync/{sid}", files={"file": ("d.bin", diffs)}); open(out,"wb").write(r.content)
chk = ic.load_from_xlsx(out,"en","UTC"); chk.evaluate(); print("saved B11 =", chk.get_formatted_cell_value(0,11,2))
print(c.post("/api/ai", json={"items": [{"sheet":0,"row":7,"col":3,"current":"","comment":"承認印の有無を確認した旨を記載"}]}).json())
