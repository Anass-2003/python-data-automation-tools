import csv

# 1. بيانات مرتبة في قائمة (مثل أعمدة وصفوف الإكسل)
data = [
    ["Name", "Balance", "Status"],
    ["Anass", 1500, "Active"],
    ["hajar", 2300, "Active"],
    ["Karim", 0, "Inactive"]
]

# 2. إنشاء ملف CSV وحفظ البيانات فيه
with open("customers.csv", "w", newline="") as file:
    writer = csv.writer(file)
    writer.writerows(data)

print("-> CSV file created successfully!")


# --- قراءة البيانات وتصفيتها ---

print("\n--- Reading Active Customers Only ---")

with open("customers.csv", "r") as file:
    reader = csv.reader(file)
    
    # تخطي السطر الأول (رؤوس الأعمدة: Name, Balance, Status)
    header = next(reader) 
    
    for row in reader:
        name = row[0]
        balance =  float(row[1])
        status = row[2]
        
        # شرط الفلترة: طباعة الحسابات النشطة فقط
        if status == "Active"and balance <=2000:
            print(f"Customer: {name} | Balance: ${balance}")


            import csv

filtered_data = [["Name", "Balance", "Status"]]  # رأس الجدول للملف الجديد

with open("customers.csv", "r") as file:
    reader = csv.reader(file)
    header = next(reader)
    
    for row in reader:
        name = row[0]
        balance = float(row[1])
        status = row[2]
        
        # حفظ الحسابات النشطة التي رصيدها 2000 أو أقل
        if status == "Active" and balance <= 2000:
            filtered_data.append([name, balance, status])

# كتابة البيانات المفلترة في ملف جديد
with open("filtered_customers.csv", "w", newline="") as new_file:
    writer = csv.writer(new_file)
    writer.writerows(filtered_data)

print("-> Filtered data saved to 'filtered_customers.csv'!")