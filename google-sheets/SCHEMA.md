# FRAME Medicine — Google Sheets Schema
# ======================================
# This is the source of truth for every tab and column.
# When building Code.gs, column indices start at 0.
# When writing to sheets, column numbers start at 1.
#
# RULES:
# - Patients.Name is the primary key across all tabs
# - Phone numbers stored as +1XXXXXXXXXX
# - Dates stored as real Date objects, never text
# - All calculations use real calendar dates (actual days in month)
# - Never use 4.333 or any fixed weeks-per-month constant
# - Medication costs sourced from Catalog tab, never hardcoded
# - Monthly overhead is a variable per-month input, not a fixed constant

## Column Index Constants (for Code.gs)

### Patients Tab
P_NAME=0, P_PREFERRED=1, P_DOB=2, P_PHONE=3, P_EMAIL=4,
P_ADDR=5, P_CITY=6, P_STATE=7, P_ZIP=8, P_SINCE=9,
P_MED=10, P_PLAN=11, P_RATE=12, P_TERM=13, P_MEMSTART=14,
P_CONTEND=15, P_CYCLES=16, P_OUTSTANDING=17,
P_CIDAY=18, P_CITIME=19, P_GLPDAY=20, P_GLPTIME=21,
P_STATUS=22, P_FOLLOWUP=23, P_NOTES=24,
P_PUSH=25, P_PUSHSUB=26, P_REFSOURCE=27, P_REFBY=28

### Medications Tab (orders start at row 15, 0-indexed row 14)
M_ORDERDATE=0, M_PATIENT=1, M_PHONE=2, M_MED=3, M_FORM=4,
M_DOSE=5, M_VIALS=6, M_DAYS=7, M_SHIPDATE=8, M_NEXTDUE=9,
M_VIALCOST=10, M_SUPPLY=11, M_SHIPPING=12, M_TOTAL=13,
M_MONTHLY=14, M_NOTES=15

### Billing Tab (internal billing tracking only)
S_PATIENT=0, S_PLAN=1, S_RATE=2, S_TERM=3, S_MEMSTART=4,
S_LASTPAY=5, S_CONTEND=6, S_CYCLES=7, S_OUTSTANDING=8,
S_STATUS=9, S_LASTSHIP=10, S_NEXTSHIP=11, S_NEXTPAYDUE=12, S_NOTES=13

### Labs Tab
L_PATIENT=0, L_ENROLL=1, L_INIT_DATE=2, L_INIT_DONE=3,
L_90_DUE=4, L_90_DONE=5, L_180_DUE=6, L_180_DONE=7,
L_ANN_DUE=8, L_ANN_DONE=9, L_NEXT_DUE=10, L_STATUS=11, L_NOTES=12

### Leads Tab
LD_NAME=0, LD_PHONE=1, LD_EMAIL=2, LD_SOURCE=3, LD_DATE=4,
LD_INTEREST=5, LD_STAGE=6, LD_ASSIGNED=7, LD_LASTCONTACT=8,
LD_NEXTFOLLOWUP=9, LD_NOTES=10, LD_CONVERTED=11,
LD_CONVERTEDDATE=12, LD_PATIENTNAME=13

### Messages Tab
MSG_TIMESTAMP=0, MSG_PATIENT=1, MSG_PHONE=2, MSG_DIRECTION=3,
MSG_TEXT=4, MSG_READ=5, MSG_SOURCE=6, MSG_CONTACTTYPE=7

### Dose History Tab
DH_DATE=0, DH_PATIENT=1, DH_MED=2, DH_OLDDOSE=3,
DH_NEWDOSE=4, DH_CHANGEDBY=5, DH_REASON=6

### Finance Tab
FIN_MONTH=0, FIN_YEAR=1, FIN_MONTHNUM=2, FIN_REVENUE=3,
FIN_MEDCOSTS=4, FIN_OVERHEAD=5, FIN_NET=6, FIN_TOM=7,
FIN_COLIN=8, FIN_LOCKED=9

### Overhead Items Tab
OH_MONTH=0, OH_YEAR=1, OH_DESC=2, OH_AMOUNT=3

## Patient Status Values
- Active
- Active - No Med
- On Hold
- Declined Refill
- INACTIVE
- Staff

## Lead Stage Values
- Inquiry
- Consultation Scheduled
- Consultation Done
- Enrolled

## Billing Status Values (internal only)
- Active
- Expired
- Past Due
- Cancelled
