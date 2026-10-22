# -*- coding: utf-8 -*-
"""FastAPI backend for polyphenol database."""
import os
import io
import json
import sqlite3
from typing import Optional
from fastapi import FastAPI, UploadFile, File
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import JSONResponse, FileResponse
from fastapi.staticfiles import StaticFiles
import openpyxl

DB_PATH = os.path.join(os.path.dirname(__file__), "polyphenol.db")
SEED_PATH = os.path.join(os.path.dirname(__file__), "seed.sql")

if not os.path.exists(DB_PATH) and os.path.exists(SEED_PATH):
    _conn = sqlite3.connect(DB_PATH)
    with open(SEED_PATH, "r", encoding="utf-8") as f:
        _conn.executescript(f.read())
    _conn.commit()
    _conn.close()
    print("Initialized database from seed.sql")

app = FastAPI(title="Polyphenol DB")
app.add_middleware(CORSMiddleware, allow_origins=["*"], allow_credentials=True, allow_methods=["*"], allow_headers=["*"])

STATIC_DIR = os.path.join(os.path.dirname(__file__), "static")
os.makedirs(STATIC_DIR, exist_ok=True)

@app.get("/")
def index():
    return FileResponse(os.path.join(STATIC_DIR, "index.html"))

def get_db():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn

@app.get("/api/stats")
def stats():
    conn = get_db()
    c = conn.cursor()
    total = c.execute("SELECT COUNT(*) FROM compounds").fetchone()[0]
    ev = c.execute("SELECT COUNT(*) FROM evidence").fetchone()[0]
    up = c.execute("SELECT COUNT(*) FROM uplc").fetchone()[0]
    ref = c.execute("SELECT COUNT(*) FROM lit_references").fetchone()[0]
    smi = c.execute("SELECT COUNT(*) FROM compounds WHERE smiles IS NOT NULL OR foodb_smiles IS NOT NULL").fetchone()[0]
    cd = {}
    for r in c.execute("SELECT primary_class, COUNT(*) n FROM compounds GROUP BY primary_class"):
        cd[r[0] or "Unknown"] = r[1]
    cr = {}
    for col, key in [("mce_phenols","MCE"),("targetmol_phenols","TargetMol"),("phenol_explorer","PhenolExplorer"),("foodb_phenolics","FooDB")]:
        cr[key] = c.execute(f"SELECT COUNT(*) FROM compounds WHERE {col}='Yes'").fetchone()[0]
    conn.close()
    return {"compounds_total":total,"evidence_total":ev,"uplc_total":up,"references_total":ref,"has_smiles":smi,"class_dist":cd,"cross_db":cr}

@app.get("/api/filters")
def filters():
    conn = get_db()
    classes = [r[0] for r in conn.execute("SELECT DISTINCT primary_class FROM compounds WHERE primary_class IS NOT NULL ORDER BY primary_class")]
    plants = set()
    for r in conn.execute("SELECT source FROM evidence WHERE source IS NOT NULL"):
        for p in str(r[0]).replace("；",";").split(";"):
            p = p.strip()
            if p: plants.add(p)
    conn.close()
    return {"classes":classes,"plants":sorted(plants)}

@app.get("/api/compounds")
def list_compounds(keyword="", primary_class="", cas="", plants="", has_smiles=False, has_evidence=False, has_uplc=False, page=1, page_size=50):
    conn = get_db()
    where, params = [], []
    if keyword:
        kw = f"%{keyword}%"
        where.append("(name_en LIKE ? OR name_cn LIKE ? OR cas LIKE ? OR formula LIKE ? OR primary_class LIKE ?)")
        params += [kw,kw,kw,kw,kw]
    if primary_class: where.append("primary_class=?"); params.append(primary_class)
    if cas: where.append("cas LIKE ?"); params.append(f"%{cas}%")
    if has_smiles: where.append("(smiles IS NOT NULL OR foodb_smiles IS NOT NULL)")
    if has_evidence: where.append("name_en IN (SELECT DISTINCT compound_name FROM evidence)")
    if has_uplc: where.append("name_en IN (SELECT DISTINCT compound_name FROM uplc)")
    if plants:
        pl = [p.strip() for p in plants.split(",") if p.strip()]
        if pl:
            ors = ["name_en IN (SELECT compound_name FROM evidence WHERE source LIKE ?)" for _ in pl]
            params += [f"%{p}%" for p in pl]
            where.append("(" + " OR ".join(ors) + ")")
    ws = ("WHERE "+" AND ".join(where)) if where else ""
    total = conn.execute(f"SELECT COUNT(*) FROM compounds {ws}", params).fetchone()[0]
    off = (page-1)*page_size
    rows = conn.execute(f"""SELECT name_en,name_cn,primary_class,secondary_class,cas,formula,
        mce_phenols,targetmol_phenols,phenol_explorer,foodb_phenolics,
        (SELECT COUNT(*) FROM evidence WHERE evidence.compound_name=compounds.name_en) ec
        FROM compounds {ws} ORDER BY name_en LIMIT ? OFFSET ?""", params+[page_size,off]).fetchall()
    conn.close()
    return {"total":total,"page":page,"page_size":page_size,"items":[dict(r) for r in rows]}

@app.get("/api/compounds/{name}")
def detail(name):
    conn = get_db()
    r = conn.execute("SELECT * FROM compounds WHERE name_en=?",[name]).fetchone()
    if not r: conn.close(); return JSONResponse({"error":"not found"},status_code=404)
    c = {k:r[k] for k in r.keys()}
    c["evidence_list"] = [dict(x) for x in conn.execute("SELECT * FROM evidence WHERE compound_name=? ORDER BY id",[name])]
    c["uplc_list"] = [dict(x) for x in conn.execute("SELECT * FROM uplc WHERE compound_name=? ORDER BY id",[name])]
    conn.close()
    return c

@app.get("/api/top-evidence")
def top_ev(n=20):
    conn = get_db()
    rows = conn.execute("SELECT compound_name, COUNT(*) c FROM evidence GROUP BY compound_name ORDER BY c DESC LIMIT ?",[n]).fetchall()
    conn.close()
    return [{"name":r[0],"count":r[1]} for r in rows]

@app.get("/api/top-plants")
def top_pl(n=20):
    conn = get_db()
    cnt = {}
    for r in conn.execute("SELECT source FROM evidence WHERE source IS NOT NULL"):
        for p in str(r[0]).replace("；",";").split(";"):
            p=p.strip()
            if p: cnt[p]=cnt.get(p,0)+1
    conn.close()
    return [{"name":k,"count":v} for k,v in sorted(cnt.items(),key=lambda x:-x[1])[:n]]

@app.post("/api/upload")
async def upload(file: UploadFile = File(...)):
    content = await file.read()
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    added = {"compounds":0,"evidence":0,"uplc":0}
    try:
        if file.filename.endswith(".json"):
            data = json.loads(content)
            for item in data.get("compounds",[]):
                if item.get("name_en"):
                    c.execute("INSERT OR REPLACE INTO compounds (name_en,name_cn,primary_class,cas,formula,smiles) VALUES (?,?,?,?,?,?)",
                        [item.get("name_en"),item.get("name_cn"),item.get("primary_class"),item.get("cas"),item.get("formula"),item.get("smiles")])
                    added["compounds"]+=1
            for item in data.get("evidence",[]):
                if item.get("化合物名称"):
                    c.execute("INSERT INTO evidence (compound_name,category,source,reference,data_source,doi,url,snippet) VALUES (?,?,?,?,?,?,?,?)",
                        [item.get("化合物名称"),item.get("类别"),item.get("来源"),item.get("参考文献"),item.get("数据来源"),item.get("DOI"),item.get("文献链接"),item.get("证据片段")])
                    added["evidence"]+=1
            for item in data.get("uplc",[]):
                if item.get("化合物名称"):
                    c.execute("INSERT INTO uplc (compound_name,formula,q1,q3,mw,ionization,level,cas,kegg_map) VALUES (?,?,?,?,?,?,?,?,?)",
                        [item.get("化合物名称"),item.get("Formula"),item.get("Q1 (Da)"),item.get("Q3 (Da)"),item.get("Molecular weight (Da)"),item.get("Ionization model"),item.get("Level"),item.get("CAS"),item.get("kegg_map")])
                    added["uplc"]+=1
        else:
            wb = openpyxl.load_workbook(io.BytesIO(content))
            for sn in wb.sheetnames:
                ws = wb[sn]
                rows = list(ws.iter_rows(values_only=True))
                if not rows: continue
                headers = [str(v).strip() if v else "" for v in rows[0]]
                for row in rows[1:]:
                    d = {headers[i]:row[i] for i in range(len(headers)) if i<len(row) and row[i] is not None}
                    if "化合物名称" in d:
                        c.execute("INSERT INTO evidence (compound_name,category,source,reference,data_source,doi,url,snippet) VALUES (?,?,?,?,?,?,?,?)",
                            [d.get("化合物名称"),d.get("类别"),d.get("来源"),d.get("参考文献"),d.get("数据来源"),d.get("DOI"),d.get("文献链接"),d.get("证据片段")])
                        added["evidence"]+=1
        conn.commit()
    except Exception as e:
        conn.rollback()
        return JSONResponse({"error":str(e)},status_code=400)
    finally:
        conn.close()
    return {"status":"ok","added":added,"filename":file.filename}

app.mount("/static", StaticFiles(directory=STATIC_DIR), name="static")

if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=int(os.environ.get("PORT",8000)))
