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

def init_schema():
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    c.execute("""CREATE TABLE IF NOT EXISTS compounds (
        name_en TEXT PRIMARY KEY, name_cn TEXT, primary_class TEXT, secondary_class TEXT,
        family TEXT, inclusion_level TEXT, cas TEXT, formula TEXT,
        representative_plants TEXT, common_sources TEXT,
        mce_phenols TEXT, targetmol_phenols TEXT, phenol_explorer TEXT,
        mce_url TEXT, targetmol_url TEXT, phenol_explorer_url TEXT,
        foodb_phenolics TEXT, foodb_public_id TEXT, foodb_evidence_level TEXT,
        foodb_superklass TEXT, foodb_klass TEXT, foodb_subklass TEXT,
        foodb_inchikey TEXT, foodb_smiles TEXT, foodb_annotation_quality TEXT,
        foodb_description TEXT, smiles TEXT, smiles_source TEXT, pubchem_title TEXT)""")
    c.execute("""CREATE TABLE IF NOT EXISTS evidence (
        id INTEGER PRIMARY KEY AUTOINCREMENT, compound_name TEXT, category TEXT,
        source TEXT, reference TEXT, data_source TEXT, platform TEXT, pdf_file TEXT,
        doi TEXT, url TEXT, snippet TEXT)""")
    c.execute("""CREATE TABLE IF NOT EXISTS uplc (
        id INTEGER PRIMARY KEY AUTOINCREMENT, compound_name TEXT, original_name TEXT,
        category TEXT, primary_class TEXT, secondary_class TEXT, formula TEXT,
        q1 TEXT, q3 TEXT, mw TEXT, ionization TEXT, level TEXT, cas TEXT,
        cpd_id TEXT, fp TEXT, ep TEXT, kegg_map TEXT)""")
    c.execute("""CREATE TABLE IF NOT EXISTS lit_references (
        id INTEGER PRIMARY KEY AUTOINCREMENT, crop TEXT, platform TEXT, query TEXT,
        title TEXT, doi TEXT, year TEXT, journal TEXT, url TEXT,
        compounds_identified TEXT, count TEXT)""")
    c.execute("CREATE INDEX IF NOT EXISTS idx_ev_name ON evidence(compound_name)")
    c.execute("CREATE INDEX IF NOT EXISTS idx_up_name ON uplc(compound_name)")
    c.execute("CREATE INDEX IF NOT EXISTS idx_comp_class ON compounds(primary_class)")
    conn.commit()
    conn.close()

if not os.path.exists(DB_PATH):
    init_schema()
    if os.path.exists(SEED_PATH):
        conn = sqlite3.connect(DB_PATH)
        with open(SEED_PATH, "r", encoding="utf-8") as f:
            conn.executescript(f.read())
        conn.commit()
        conn.close()
        print("Initialized database from seed.sql")
    else:
        print("Empty database created. Upload xlsx files via web UI.")

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


def _safe(v):
    if v is None: return None
    s = str(v).strip()
    return s if s else None


@app.post("/api/upload")
async def upload(file: UploadFile = File(...)):
    content = await file.read()
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    added = {"compounds":0,"evidence":0,"uplc":0,"lit_references":0}

    try:
        wb = openpyxl.load_workbook(io.BytesIO(content), read_only=True, data_only=True)
        for sn in wb.sheetnames:
            ws = wb[sn]
            rows_iter = ws.iter_rows(values_only=True)
            try:
                header = next(rows_iter)
            except StopIteration:
                continue
            headers = [str(v).strip() if v else "" for v in header]
            col = {h: i for i, h in enumerate(headers) if h}

            if sn in ("Master_complete_only", "Master_strict_complete") or "name_en" in col:
                for row in rows_iter:
                    name = _safe(row[col["name_en"]]) if "name_en" in col else None
                    if not name:
                        continue
                    c.execute("""INSERT OR REPLACE INTO compounds
                        (name_en,name_cn,primary_class,secondary_class,family,inclusion_level,cas,formula,
                         representative_plants,common_sources,mce_phenols,targetmol_phenols,phenol_explorer,
                         mce_url,targetmol_url,phenol_explorer_url,foodb_phenolics,foodb_public_id,
                         foodb_evidence_level,foodb_superklass,foodb_klass,foodb_subklass,foodb_inchikey,
                         foodb_smiles,foodb_annotation_quality,foodb_description,smiles,smiles_source,pubchem_title)
                        VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
                        [name,
                         _safe(row[col["name_cn"]]) if "name_cn" in col else None,
                         _safe(row[col["primary_class"]]) if "primary_class" in col else None,
                         _safe(row[col["secondary_class"]]) if "secondary_class" in col else None,
                         _safe(row[col["family"]]) if "family" in col else None,
                         _safe(row[col["inclusion_level"]]) if "inclusion_level" in col else None,
                         _safe(row[col["cas"]]) if "cas" in col else None,
                         _safe(row[col["formula"]]) if "formula" in col else None,
                         _safe(row[col["representative_plants"]]) if "representative_plants" in col else None,
                         _safe(row[col["common_sources"]]) if "common_sources" in col else None,
                         _safe(row[col["mce_phenols"]]) if "mce_phenols" in col else None,
                         _safe(row[col["targetmol_polyphenols"]]) if "targetmol_polyphenols" in col else (_safe(row[col["targetmol_phenols"]]) if "targetmol_phenols" in col else None),
                         _safe(row[col["phenol_explorer"]]) if "phenol_explorer" in col else None,
                         _safe(row[col["mce_url"]]) if "mce_url" in col else None,
                         _safe(row[col["targetmol_url"]]) if "targetmol_url" in col else None,
                         _safe(row[col["phenol_explorer_url"]]) if "phenol_explorer_url" in col else None,
                         _safe(row[col["foodb_phenolics"]]) if "foodb_phenolics" in col else None,
                         _safe(row[col["foodb_public_id"]]) if "foodb_public_id" in col else None,
                         _safe(row[col["foodb_evidence_level"]]) if "foodb_evidence_level" in col else None,
                         _safe(row[col["foodb_superklass"]]) if "foodb_superklass" in col else None,
                         _safe(row[col["foodb_klass"]]) if "foodb_klass" in col else None,
                         _safe(row[col["foodb_subklass"]]) if "foodb_subklass" in col else None,
                         _safe(row[col["foodb_inchikey"]]) if "foodb_inchikey" in col else None,
                         _safe(row[col["foodb_smiles"]]) if "foodb_smiles" in col else None,
                         _safe(row[col["foodb_annotation_quality"]]) if "foodb_annotation_quality" in col else None,
                         _safe(row[col["foodb_description"]]) if "foodb_description" in col else None,
                         _safe(row[col["smiles"]]) if "smiles" in col else None,
                         _safe(row[col["smiles_source"]]) if "smiles_source" in col else None,
                         _safe(row[col["pubchem_title"]]) if "pubchem_title" in col else None])
                    added["compounds"] += 1

            elif sn == "补充明细" or ("化合物名称" in col and "证据片段" in col):
                for row in rows_iter:
                    name = _safe(row[col["化合物名称"]]) if "化合物名称" in col else None
                    if not name:
                        continue
                    c.execute("""INSERT INTO evidence
                        (compound_name,category,source,reference,data_source,platform,pdf_file,doi,url,snippet)
                        VALUES (?,?,?,?,?,?,?,?,?,?)""",
                        [name,
                         _safe(row[col["类别"]]) if "类别" in col else None,
                         _safe(row[col["来源"]]) if "来源" in col else None,
                         _safe(row[col["参考文献"]]) if "参考文献" in col else None,
                         _safe(row[col["数据来源"]]) if "数据来源" in col else None,
                         _safe(row[col["来源平台"]]) if "来源平台" in col else None,
                         _safe(row[col["PDF文件"]]) if "PDF文件" in col else None,
                         _safe(row[col["DOI"]]) if "DOI" in col else None,
                         _safe(row[col["文献链接"]]) if "文献链接" in col else None,
                         _safe(row[col["证据片段"]]) if "证据片段" in col else None])
                    added["evidence"] += 1

            elif sn == "菜籽UPLC筛选" or ("Q1 (Da)" in col and "Ionization model" in col):
                for row in rows_iter:
                    name = _safe(row[col["化合物名称"]]) if "化合物名称" in col else None
                    if not name:
                        continue
                    c.execute("""INSERT INTO uplc
                        (compound_name,original_name,category,primary_class,secondary_class,formula,
                         q1,q3,mw,ionization,level,cas,cpd_id,fp,ep,kegg_map)
                        VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
                        [name,
                         _safe(row[col["原始化合物名称"]]) if "原始化合物名称" in col else None,
                         _safe(row[col["类别"]]) if "类别" in col else None,
                         _safe(row[col["一级分类"]]) if "一级分类" in col else None,
                         _safe(row[col["二级分类"]]) if "二级分类" in col else None,
                         _safe(row[col["Formula"]]) if "Formula" in col else None,
                         _safe(row[col["Q1 (Da)"]]) if "Q1 (Da)" in col else None,
                         _safe(row[col["Q3 (Da)"]]) if "Q3 (Da)" in col else None,
                         _safe(row[col["Molecular weight (Da)"]]) if "Molecular weight (Da)" in col else None,
                         _safe(row[col["Ionization model"]]) if "Ionization model" in col else None,
                         _safe(row[col["Level"]]) if "Level" in col else None,
                         _safe(row[col["CAS"]]) if "CAS" in col else None,
                         _safe(row[col["cpd_ID"]]) if "cpd_ID" in col else None,
                         _safe(row[col["FP"]]) if "FP" in col else None,
                         _safe(row[col["EP"]]) if "EP" in col else None,
                         _safe(row[col["kegg_map"]]) if "kegg_map" in col else None])
                    added["uplc"] += 1

            elif sn == "在线化学实体记录" or "检索作物" in col:
                for row in rows_iter:
                    title = _safe(row[col["题名"]]) if "题名" in col else None
                    if not title:
                        continue
                    c.execute("""INSERT INTO lit_references
                        (crop,platform,query,title,doi,year,journal,url,compounds_identified,count)
                        VALUES (?,?,?,?,?,?,?,?,?,?)""",
                        [_safe(row[col["检索作物"]]) if "检索作物" in col else None,
                         _safe(row[col["来源平台"]]) if "来源平台" in col else None,
                         _safe(row[col["检索词"]]) if "检索词" in col else None,
                         title,
                         _safe(row[col["DOI"]]) if "DOI" in col else None,
                         _safe(row[col["年份"]]) if "年份" in col else None,
                         _safe(row[col["期刊"]]) if "期刊" in col else None,
                         _safe(row[col["文献链接"]]) if "文献链接" in col else None,
                         _safe(row[col["识别化合物"]]) if "识别化合物" in col else None,
                         _safe(row[col["识别数量"]]) if "识别数量" in col else None])
                    added["lit_references"] += 1

        wb.close()
        conn.commit()
    except Exception as e:
        conn.rollback()
        import traceback
        return JSONResponse({"error": str(e), "detail": traceback.format_exc()}, status_code=400)
    finally:
        conn.close()

    return {"status":"ok","added":added,"filename":file.filename}


app.mount("/static", StaticFiles(directory=STATIC_DIR), name="static")

if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=int(os.environ.get("PORT",8000)))
