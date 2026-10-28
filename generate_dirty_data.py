"""
بنولّد هنا ملف بيانات عملاء "وسخ" بشكل مقصود (كله بيانات تجريبية).
الملف اللي بيطلع: dirty_customers.csv
"""
import random
import numpy as np
import pandas as pd
from datetime import datetime, timedelta

random.seed(42)
np.random.seed(42)

N = 3000

first_m = ["محمد", "أحمد", "محمود", "مصطفى", "عمر", "علي", "حسن", "إبراهيم", "خالد", "يوسف",
           "كريم", "طارق", "هاني", "سامح", "وليد", "ياسر", "عبدالله", "مروان", "تامر", "شريف"]
first_f = ["فاطمة", "نور", "سارة", "منى", "هدى", "ياسمين", "آية", "دينا", "رنا", "مريم",
           "نادية", "سلمى", "إيمان", "هبة", "أسماء", "ريم", "جميلة", "داليا", "شيماء", "لمياء"]
last = ["السيد", "عبدالعزيز", "الشريف", "حسين", "منصور", "فؤاد", "رمضان", "إسماعيل", "الجمل",
        "عبدالرحمن", "صالح", "النجار", "الحداد", "عثمان", "زكريا", "سليمان", "البنا", "جابر"]

cities_ok = ["القاهرة", "الجيزة", "الإسكندرية", "المنصورة", "طنطا", "أسيوط",
             "الأقصر", "أسوان", "بورسعيد", "السويس"]
city_typos = {
    "القاهرة": ["القاهره", "cairo", "Cairo", "القاهرة ", "القاهرا"],
    "الجيزة": ["الجيزه", "Giza", "giza", "الجيزة "],
    "الإسكندرية": ["الاسكندرية", "اسكندريه", "الأسكندرية", "Alexandria", "alex"],
    "المنصورة": ["المنصوره", "mansoura", "المنصورة "],
    "طنطا": ["طنطه", "Tanta", "tanta"],
    "أسيوط": ["اسيوط", "Assiut", "أسيوط "],
    "الأقصر": ["الاقصر", "Luxor", "luxor"],
    "أسوان": ["اسوان", "Aswan"],
    "بورسعيد": ["بور سعيد", "Port Said", "بورسعيد "],
    "السويس": ["السويس ", "Suez", "suez"],
}
weights = [0.28, 0.17, 0.17, 0.07, 0.06, 0.05, 0.04, 0.04, 0.06, 0.06]

rows = []
for i in range(1, N + 1):
    gender = random.choice(["M", "F"])
    first = random.choice(first_m if gender == "M" else first_f)
    name = f"{first} {random.choice(last)}"
    city = np.random.choice(cities_ok, p=weights)
    prefix = random.choice(["10", "11", "12", "15"])
    phone_core = prefix + "".join(random.choices("0123456789", k=8))
    email_user = f"user{random.randint(1000, 99999)}"
    email = f"{email_user}@{random.choice(['gmail.com', 'yahoo.com', 'hotmail.com', 'outlook.com'])}"
    signup = datetime(2023, 1, 1) + timedelta(days=random.randint(0, 900))
    age = int(np.clip(np.random.normal(33, 10), 18, 65))
    spent = round(float(np.random.gamma(2.2, 900)), 2)
    rows.append(dict(customer_id=i, name=name, phone_core=phone_core, email=email, city=city,
                     signup=signup, age=age, gender=gender, total_spent=spent))

df = pd.DataFrame(rows)

# ---------- بنحط الغلطات في الداتا ----------
def idx(frac):
    return df.sample(frac=frac, random_state=random.randint(0, 10_000)).index

# 1) الأسماء: مسافات زيادة
for i in idx(0.12):
    df.loc[i, "name"] = random.choice([f"  {df.loc[i,'name']}", f"{df.loc[i,'name']}  ", df.loc[i, "name"].replace(" ", "   ")])

# 2) الموبايل بصيغ مختلفة
def phone_fmt(core):
    return random.choice([
        "0" + core, "+20" + core, "0020" + core, "0" + core[:2] + "-" + core[2:6] + "-" + core[6:],
        core, "0" + core[:3] + " " + core[3:], "(+20) " + core,
    ])
df["phone"] = df["phone_core"].apply(phone_fmt)
for i in idx(0.015):  # أرقام مش صح
    df.loc[i, "phone"] = random.choice(["12345", "0100000", "غير متوفر", "N/A", "01x2345678"])

# 3) التواريخ بصيغ مختلفة
def date_fmt(d):
    return d.strftime(random.choice(["%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y", "%d %b %Y", "%Y/%m/%d"]))
df["signup_date"] = df["signup"].apply(date_fmt)
for i in idx(0.01):
    df.loc[i, "signup_date"] = random.choice(["31/02/2024", "تاريخ غير معروف", "0000-00-00"])

# 4) الإيميل: حروف كابيتال وإيميلات مش صح
for i in idx(0.10):
    df.loc[i, "email"] = df.loc[i, "email"].upper()
for i in idx(0.02):
    df.loc[i, "email"] = random.choice(["user@@gmail.com", "abc.com", "test@", "no email"])

# 5) المدن: غلطات إملاء وأسماء بالإنجليزي
df["city"] = df["city"].astype(object)
for i in idx(0.30):
    c = df.loc[i, "city"]
    df.loc[i, "city"] = random.choice(city_typos[c])

# 6) الجنس بصيغ مختلفة
gmap = {"M": ["ذكر", "male", "M", "Male", "ذكر "], "F": ["أنثى", "انثى", "female", "F", "Female"]}
df["gender"] = df["gender"].apply(lambda g: random.choice(gmap[g]) if random.random() < 0.45 else ("ذكر" if g == "M" else "أنثى"))

# 7) قيم غريبة (شاذة)
for i in idx(0.006):
    df.loc[i, "age"] = random.choice([250, 0, -5, 999, 150])
for i in idx(0.006):
    df.loc[i, "total_spent"] = random.choice([-500.0, -1.0, 9_999_999.0, 1_000_000.0])

# 8) قيم فاضية
for col, frac in [("email", 0.06), ("age", 0.05), ("city", 0.04), ("gender", 0.03), ("phone", 0.03)]:
    df.loc[idx(frac), col] = np.nan

# 9) تكرارات: نسخ طبق الأصل ونسخ لنفس الشخص بس بصيغة مختلفة
exact = df.sample(n=90, random_state=1)
near = df.sample(n=60, random_state=2).copy()
near["name"] = near["name"].astype(str).str.strip().apply(lambda s: s + " ")
near["email"] = near["email"].astype(str).str.lower()
near["customer_id"] = near["customer_id"] + 10_000
dirty = pd.concat([df, exact, near], ignore_index=True).sample(frac=1, random_state=3).reset_index(drop=True)

out = dirty[["customer_id", "name", "phone", "email", "city", "signup_date", "age", "gender", "total_spent"]]
out.to_csv("dirty_customers.csv", index=False, encoding="utf-8-sig")
print(f"تم توليد {len(out)} صف في dirty_customers.csv")
