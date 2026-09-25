# -*- coding: utf-8 -*-
from pathlib import Path
import datetime as dt, json, re, sys
import openpyxl
BASE=Path(__file__).resolve().parent
EXCEL=BASE/"DATA GRUAS.xlsx"; TEMPLATE=BASE/"index.html"; OUTPUT=BASE/"index.html"
def date(v):
    if isinstance(v,dt.datetime): return v.date()
    if isinstance(v,dt.date): return v
    return None
def num(v):
    try:return float(v or 0)
    except:return 0.0
def build():
    wb=openpyxl.load_workbook(EXCEL,data_only=True,read_only=True)
    reqs={"DATA","DM_USAGE",'Curva "S"'}; miss=reqs-set(wb.sheetnames)
    if miss: raise ValueError("Faltan hojas: "+", ".join(sorted(miss)))
    ws=wb["DATA"]; head=next(ws.iter_rows(min_row=1,max_row=1,values_only=True))
    h={str(v).strip():i for i,v in enumerate(head) if v is not None and str(v).strip()}
    req=["Fecha","ESTADO","Detalle de Actividad","Equipo","Horas Efectivas o metrados","HDNT","HND","HNP"]
    miss=[x for x in req if x not in h]
    if miss: raise ValueError("Faltan columnas en DATA: "+", ".join(miss))
    dates=set(); data={}; production={}
    for r in ws.iter_rows(min_row=2,values_only=True):
        d=date(r[h["Fecha"]])
        if not d: continue
        ds=d.isoformat(); dates.add(ds); st=str(r[h["ESTADO"]] or "").strip().upper()
        eq=str(r[h["Equipo"]] or "Sin equipo").strip(); act=str(r[h["Detalle de Actividad"]] or "Sin detalle").strip()
        obs=str(r[h["Observaciones"]] or "").strip() if "Observaciones" in h else ""
        if st in ("STAND_BY","NO_DISPONIBLE","NO_PROGRAMADO"):
            col={"STAND_BY":"HDNT","NO_DISPONIBLE":"HND","NO_PROGRAMADO":"HNP"}[st]
            data.setdefault(ds,{}).setdefault(st,[]).append({"equipo":eq,"actividad":act,"observaciones":obs,"horas":num(r[h[col]])})
        elif st=="PRODUCCIÓN": production[ds]=production.get(ds,0)+num(r[h["Horas Efectivas o metrados"]])
    dates=sorted(dates)
    for d in dates:data.setdefault(d,{})
    ws=wb["DM_USAGE"]; dm={}; usage={}
    for r in ws.iter_rows(min_row=2,values_only=True):
        d=date(r[1]); eq=str(r[3] or "").strip()
        if not d or not eq: continue
        ds=d.isoformat(); dm.setdefault(ds,{})[eq]=num(r[4]); usage.setdefault(ds,{})[eq]=num(r[5])
    ws=wb['Curva "S"']; series=[]; last=dt.date.fromisoformat(dates[-1])
    for c in range(3,ws.max_column+1):
        d=date(ws.cell(29,c).value)
        if not d: continue
        vals=[ws.cell(r,c).value for r in (31,35,36,37)]
        if all(v is None for v in vals): continue
        series.append({"date":d.isoformat(),"target":num(vals[0]),"projected":d>last,
                       "Camión Grúa 23 TN":num(vals[1]),"Grua Semitrailer 20 TN":num(vals[2]),"Tadano 60 TN":num(vals[3])})
    end12=next((x["date"] for x in series if dt.date.fromisoformat(x["date"]).day==12 and x["date"]>=dates[-1]),None)
    if end12: series=[x for x in series if x["date"]<=end12]
    cd=[x["date"] for x in series]
    curva={"dates":cd,"series":series,"equipment":["Camión Grúa 23 TN","Grua Semitrailer 20 TN","Tadano 60 TN"],
            "defaultStart":dates[-1] if dates[-1] in cd else (cd[0] if cd else ""),
            "defaultEnd":next((x["date"] for x in reversed(series) if dt.date.fromisoformat(x["date"]).day==12),cd[-1] if cd else ""),
            "lastActual":dates[-1],"source":'Curva "S"',"projectionEndpointDay":12}
    return {"dates":dates,"data":data,"production":production,"source":"DATA","range":[dates[0],dates[-1]],"dmUsage":dm,"usage":usage,"dmTarget":0.92},curva
def main():
    if not EXCEL.exists(): raise FileNotFoundError("No se encontró DATA GRUAS.xlsx")
    p,c=build(); html=TEMPLATE.read_text(encoding="utf-8")
    html=re.sub(r'const PAYLOAD = (\{.*?\});\s*\n','const PAYLOAD = '+json.dumps(p,ensure_ascii=False,separators=(",",":"))+';\n',html,count=1,flags=re.S)
    html=re.sub(r'const CURVA_DATA = (\{.*?\});\s*\n','const CURVA_DATA = '+json.dumps(c,ensure_ascii=False,separators=(",",":"))+';\n',html,count=1,flags=re.S)
    OUTPUT.write_text(html,encoding="utf-8")
    print("Dashboard actualizado:",OUTPUT.name); print("Última fecha:",p["dates"][-1])
if __name__=="__main__":
    try: main()
    except Exception as e: print("ERROR:",e); input("Presiona ENTER para cerrar..."); sys.exit(1)
