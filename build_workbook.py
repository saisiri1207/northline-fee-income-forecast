"""Build Northline_Fee_Income_Forecast.xlsx (fictional bank operations style sample data).
Yellow cells are inputs, black cells are formulas. Python mirrors the formulas to print results."""
import random
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font
from openpyxl.utils import get_column_letter as L
random.seed(5)
Y = PatternFill("solid", fgColor="FFF2CC"); H = PatternFill("solid", fgColor="1F3864"); HF = Font(bold=True, color="FFFFFF")
mn = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"]
season = [0.92,0.90,1.02,1.00,1.03,1.04,1.01,1.02,0.99,1.01,1.00,1.06]
# fee type, base monthly volume 2024, fee per txn 2024, 2025 rate, 2026 rate, yearly volume growth
prods = [("Wire transfer fees",42000,18.0,18.0,20.0,0.06),("ACH origination fees",610000,0.30,0.30,0.33,0.09),
         ("Treasury service charges",9500,55.0,57.0,62.0,0.04),("Card interchange",1450000,0.42,0.43,0.45,0.08)]
periods = [(y,m) for y in (2024,2025) for m in range(12)] + [(2026,m) for m in range(6)]
wb = Workbook(); ws = wb.active; ws.title = "README"
for i,t in enumerate(["Northline Fee Income Forecast","Northline Bank Ops is a fictional sample. Figures are sample data, not from any real bank.",
  "Yellow cells are inputs. Black cells are formulas.","Actuals: monthly volume and fee per transaction by fee type, Jan 2024 to Jun 2026.",
  "Drivers: volume growth and latest fee rate per fee type.","Forecast: Jan to Jun 2026 naive (same month last year) vs driver based forecast, with accuracy (1 minus MAPE)."],1):
    ws.cell(i,1,t)
ws["A1"].font = Font(bold=True, size=14)
a = wb.create_sheet("Actuals")
for j,h in enumerate(["Month","Fee type","Volume","Fee per txn","Fee income"],1): c=a.cell(1,j,h); c.fill=H; c.font=HF
act = {}; r = 2; rowof = {}
for (y,m) in periods:
    for p,v0,r24,r25,r26,g in prods:
        vol = round(v0*(1+g)**(y-2024)*season[m]*random.uniform(0.90,1.10))
        rate = {2024:r24,2025:r25,2026:r26}[y]
        a.cell(r,1,f"{mn[m]} {y}"); a.cell(r,2,p)
        c=a.cell(r,3,vol); c.fill=Y; c.number_format='#,##0'
        c=a.cell(r,4,rate); c.fill=Y; c.number_format='0.00'
        a.cell(r,5,f"=C{r}*D{r}").number_format='#,##0'
        act[(y,m,p)] = (vol,rate); rowof[(y,m,p)] = r; r += 1
d = wb.create_sheet("Drivers")
for j,h in enumerate(["Fee type","Volume 2024","Volume 2025","Volume growth","Latest fee per txn"],1): c=d.cell(1,j,h); c.fill=H; c.font=HF
drv = {}
for i,(p,*_ ) in enumerate(prods,2):
    d.cell(i,1,p)
    d.cell(i,2,f'=SUMIFS(Actuals!C:C,Actuals!B:B,A{i},Actuals!A:A,"*2024")').number_format='#,##0'
    d.cell(i,3,f'=SUMIFS(Actuals!C:C,Actuals!B:B,A{i},Actuals!A:A,"*2025")').number_format='#,##0'
    d.cell(i,4,f"=C{i}/B{i}-1").number_format='0.0%'
    c=d.cell(i,5,prods[i-2][4]); c.fill=Y  # announced 2026 fee schedule
    v24=sum(act[(2024,m,p)][0] for m in range(12)); v25=sum(act[(2025,m,p)][0] for m in range(12))
    drv[p]=(v25/v24-1, prods[i-2][4])
f = wb.create_sheet("Forecast")
hdr=["Month","Fee type","Actual","Naive forecast","Driver forecast","Naive abs % error","Driver abs % error"]
for j,h in enumerate(hdr,1): c=f.cell(1,j,h); c.fill=H; c.font=HF
r=2; ne=[]; de=[]; tot={"a":0,"n":0,"d":0}
for m in range(6):
    for i,(p,*_) in enumerate(prods,2):
        ra=rowof[(2026,m,p)]; rl=rowof[(2025,m,p)]
        f.cell(r,1,f"{mn[m]} 2026"); f.cell(r,2,p)
        f.cell(r,3,f"=Actuals!E{ra}"); f.cell(r,4,f"=Actuals!E{rl}")
        f.cell(r,5,f"=Actuals!C{rl}*(1+Drivers!D{i})*Drivers!E{i}")
        f.cell(r,6,f"=ABS(D{r}-C{r})/C{r}").number_format='0.0%'; f.cell(r,7,f"=ABS(E{r}-C{r})/C{r}").number_format='0.0%'
        for k in (3,4,5): f.cell(r,k).number_format='#,##0'
        A=act[(2026,m,p)][0]*act[(2026,m,p)][1]; N=act[(2025,m,p)][0]*act[(2025,m,p)][1]
        D=act[(2025,m,p)][0]*(1+drv[p][0])*drv[p][1]
        ne.append(abs(N-A)/A); de.append(abs(D-A)/A); tot["a"]+=A; tot["n"]+=N; tot["d"]+=D; r+=1
last=r-1
f.cell(r+1,1,"Accuracy (1 minus MAPE)"); f.cell(r+1,6,f"=1-AVERAGE(F2:F{last})").number_format='0.0%'; f.cell(r+1,7,f"=1-AVERAGE(G2:G{last})").number_format='0.0%'
f.cell(r+2,1,"Jan to Jun 2026 total"); f.cell(r+2,3,f"=SUM(C2:C{last})"); f.cell(r+2,4,f"=SUM(D2:D{last})"); f.cell(r+2,5,f"=SUM(E2:E{last})")
for k in (3,4,5): f.cell(r+2,k).number_format='#,##0'
for w_ in wb.worksheets:
    w_.column_dimensions['A'].width=26; w_.column_dimensions['B'].width=26
    for k in range(3,8): w_.column_dimensions[L(k)].width=18
wb.save("Northline_Fee_Income_Forecast.xlsx")
print(f"Naive accuracy {1-sum(ne)/len(ne):.1%}  Driver accuracy {1-sum(de)/len(de):.1%}")
print(f"H1 2026 actual {tot['a']:,.0f}  naive {tot['n']:,.0f}  driver {tot['d']:,.0f}")
fy25=sum(act[(2025,m,p)][0]*act[(2025,m,p)][1] for m in range(12) for p,*_ in prods); print(f"FY2025 fee income {fy25:,.0f}")
for p in drv: print(p, f"growth {drv[p][0]:.1%}")
