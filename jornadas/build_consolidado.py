#!/usr/bin/env python3
"""python3 -I build_v2.py datos.json salida.xlsx
unidad: {unidad, fecha_inicio, fecha_termino (YYYY-MM-DD|null), jornadas:[{planeados,intervenidos,observaciones}],
 pacientes:[{nombre,paterno,materno,nss,cirugia,fecha_prog,turno_prog,realizo,fecha_iqx,turno_iqx,obs}]}"""
import json, sys, datetime as dt
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
ORDEN=["HGZ 1","HGZ 2","HGZ 6","HGZ 50","HGSZ 4","HGSZ 9"]
VERDE,CLARO="1F5C4A","E8F1EE"
thin=Side(style="thin",color="BFC9C5"); BORDE=Border(left=thin,right=thin,top=thin,bottom=thin)
HFONT=Font(bold=True,color="FFFFFF",name="Calibri",size=11); HFILL=PatternFill("solid",fgColor=VERDE)
def header(ws,row,titles):
    for i,t in enumerate(titles,1):
        c=ws.cell(row,i,t); c.font,c.fill,c.border=HFONT,HFILL,BORDE
        c.alignment=Alignment(horizontal="center",vertical="center",wrap_text=True)
    ws.row_dimensions[row].height=32
def D(s):
    if not s: return None
    try: return dt.datetime.strptime(s,"%Y-%m-%d")
    except Exception: return s
def build(datos,salida):
    por={u["unidad"]:u for u in datos["unidades"]}
    wb=Workbook(); rs=wb.active; rs.title="Resumen por unidad"
    rs["A1"]="Acción 10. Jornadas quirúrgicas extraordinarias — Resumen por unidad"; rs["A1"].font=Font(bold=True,size=14,color=VERDE)
    header(rs,3,["Unidad","Fecha inicio","Fecha término","Pacientes planeados","Pacientes intervenidos","% Logro","Pacientes en censo nominal","Jornadas registradas"])
    r0=4
    for i,u in enumerate(ORDEN):
        r=r0+i; d=por.get(u,{})
        rs.cell(r,1,u); rs.cell(r,2,D(d.get("fecha_inicio"))); rs.cell(r,3,D(d.get("fecha_termino")))
        rs.cell(r,4,f'=IF(COUNTIFS(Jornadas!$A:$A,A{r},Jornadas!$C:$C,"<>")=0,"",SUMIFS(Jornadas!$C:$C,Jornadas!$A:$A,A{r}))')
        rs.cell(r,5,f'=IF(COUNTIFS(Jornadas!$A:$A,A{r},Jornadas!$D:$D,"<>")=0,"",SUMIFS(Jornadas!$D:$D,Jornadas!$A:$A,A{r}))')
        rs.cell(r,6,f'=IF(AND(ISNUMBER(D{r}),D{r}>0,ISNUMBER(E{r})),E{r}/D{r},"")').number_format="0.0%"
        rs.cell(r,7,f'=COUNTIFS(Consolidado!$A:$A,A{r},Consolidado!$E:$E,"<>")')
        rs.cell(r,8,f'=COUNTIFS(Jornadas!$A:$A,A{r})')
        for j in range(1,9):
            c=rs.cell(r,j); c.border=BORDE
            if j in (2,3): c.number_format="dd/mm/yyyy"; c.alignment=Alignment(horizontal="center")
            if j>=4: c.alignment=Alignment(horizontal="center")
    rt=r0+len(ORDEN); rs.cell(rt,1,"TOTAL")
    rs.cell(rt,4,f"=SUM(D{r0}:D{rt-1})"); rs.cell(rt,5,f"=SUM(E{r0}:E{rt-1})")
    rs.cell(rt,6,f'=IF(D{rt}>0,E{rt}/D{rt},"")').number_format="0.0%"; rs.cell(rt,7,f"=SUM(G{r0}:G{rt-1})"); rs.cell(rt,8,f"=SUM(H{r0}:H{rt-1})")
    for j in range(1,9):
        c=rs.cell(rt,j); c.font=Font(bold=True); c.fill=PatternFill("solid",fgColor=CLARO); c.border=BORDE
        if j>=4: c.alignment=Alignment(horizontal="center")
    for col,w in zip("ABCDEFGH",[12,14,14,16,16,11,20,14]): rs.column_dimensions[col].width=w
    rs.freeze_panes="A4"
    # Jornadas
    js=wb.create_sheet("Jornadas")
    header(js,1,["Unidad","Jornada","Pacientes planeados","Pacientes intervenidos","% Logro","Observaciones"])
    r=2
    for u in ORDEN:
        for k,j in enumerate(por.get(u,{}).get("jornadas") or [],1):
            js.cell(r,1,u); js.cell(r,2,k); js.cell(r,3,j.get("planeados")); js.cell(r,4,j.get("intervenidos"))
            js.cell(r,5,f'=IF(AND(ISNUMBER(C{r}),C{r}>0,ISNUMBER(D{r})),D{r}/C{r},"")').number_format="0.0%"
            js.cell(r,6,j.get("observaciones"))
            for c in range(1,7):
                x=js.cell(r,c); x.border=BORDE
                if c in (2,3,4,5): x.alignment=Alignment(horizontal="center")
            r+=1
    for col,w in zip("ABCDEF",[11,9,16,16,10,44]): js.column_dimensions[col].width=w
    js.freeze_panes="A2"
    # Consolidado
    cs=wb.create_sheet("Consolidado",0)
    header(cs,1,["Unidad","Pacientes planeados","Pacientes intervenidos","% Logro","Nombre","Apellido Paterno","Apellido Materno","Número de Seguro Social (NSS)","Tipo de cirugía","Fecha de programación","Turno programado","¿Se realizó?","Fecha de IQx","Turno de IQx","Observaciones"])
    r=2
    for u in ORDEN:
        for p in (por.get(u,{}).get("pacientes") or [None]):
            cs.cell(r,1,u)
            cs.cell(r,2,f"=VLOOKUP($A{r},'Resumen por unidad'!$A:$F,4,FALSE)")
            cs.cell(r,3,f"=VLOOKUP($A{r},'Resumen por unidad'!$A:$F,5,FALSE)")
            cs.cell(r,4,f"=VLOOKUP($A{r},'Resumen por unidad'!$A:$F,6,FALSE)").number_format="0.0%"
            if p:
                for c,k in ((5,"nombre"),(6,"paterno"),(7,"materno"),(9,"cirugia"),(11,"turno_prog"),(12,"realizo"),(14,"turno_iqx"),(15,"obs")):
                    cs.cell(r,c,p.get(k) or None)
                n=cs.cell(r,8,p.get("nss") or None); n.number_format="@"
                for c,k in ((10,"fecha_prog"),(13,"fecha_iqx")):
                    x=cs.cell(r,c,D(p.get(k))); x.number_format="dd/mm/yyyy"
            for c in range(1,16):
                x=cs.cell(r,c); x.border=BORDE
                if c in (2,3,4,10,11,12,13,14): x.alignment=Alignment(horizontal="center")
            r+=1
    for col,w in zip("ABCDEFGHIJKLMNO",[11,14,14,10,24,20,20,26,30,16,14,12,14,14,36]): cs.column_dimensions[col].width=w
    cs.freeze_panes="A2"; cs.auto_filter.ref=f"A1:O{r-1}"; wb.save(salida)
if __name__=="__main__": build(json.load(open(sys.argv[1],encoding="utf-8")),sys.argv[2])
