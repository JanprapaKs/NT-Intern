import pandas as pd


# อ่านข้อมูลจาก Excel และ CSV
toolstaff = pd.read_csv(r'Toolstaff/Staff_v2_active_dataExport_all_page_20250115.csv')
# แสดงชื่อคอลัมน์ทั้งหมดใน toolstaff

staff = pd.read_excel(r'Toolstaff/Staff_v2_active_dataExport_all_page_202406.xlsx')

structure = pd.read_excel(r'C:\Users\NT\Desktop\Python\Toolstaff\Structure NT 27 Dec 2024_update20241212.xlsx',
    skiprows=1)  # ข้ามแถวแรกสุด



# ทำความสะอาดข้อมูล
toolstaff['ส่วนงาน'] = toolstaff['ส่วนงาน'].str.strip().str.lower()
structure['อักษรย่อ'] = structure['อักษรย่อ'].str.strip().str.lower()
structure['ชื่อตามโครงสร้าง NT'] = structure['ชื่อตามโครงสร้าง NT '].str.strip()

# ลบช่องว่าง
structure['อักษรย่อ'] = structure['อักษรย่อ'].str.replace(' ', '')
toolstaff['ส่วนงาน'] = toolstaff['ส่วนงาน'].str.replace(' ', '')
staff['ชื่อ-อังกฤษ'] = staff['ชื่อ-อังกฤษ'].str.replace(' ','')
staff['นามสกุล-อังกฤษ'] = staff['นามสกุล-อังกฤษ'].str.replace(' ','')


# เลือกเฉพาะคอลัมน์ที่ต้องการ
filtered_data = staff[['รหัสพนักงาน', 'ชื่อ-อังกฤษ', 'นามสกุล-อังกฤษ','ตำแหน่ง']]

# เลือกเฉพาะคอลัมน์ 'อักษรย่อ' และ 'ชื่อตามโครงสร้าง NT'
filtered_structure = structure[['อักษรย่อ', 'ชื่อตามโครงสร้าง NT ','สายงาน','ฝ่าย','กลุ่มงาน']]



# สร้าง Mapping ระหว่าง 'อักษรย่อ' และ 'ชื่อตามโครงสร้าง NT'
mapping_dict = dict(zip(filtered_structure['อักษรย่อ'], filtered_structure['ชื่อตามโครงสร้าง NT ']))

# เติมข้อมูล 'ชื่อเต็มส่วนงาน' โดยใช้ Mapping
toolstaff['ชื่อเต็มส่วนงาน'] = toolstaff['ส่วนงาน'].map(mapping_dict)


# รวมข้อมูลโดยใช้ 'รหัสพนักงาน' เป็นตัวเชื่อม
merged_data = pd.merge(toolstaff,filtered_data,
                    on='รหัสพนักงาน',  # ระบุคอลัมน์ที่ใช้ในการเชื่อม
                    how='left'   # ใช้การรวมแบบ 'left join' เพื่อคงข้อมูลจาก CSV ไว้ทั้งหมด
)
print(merged_data.columns)

#รวมข้อมูลระหว่าง merged_data และ filtered_structure 
merged_data_structure = pd.merge(merged_data,filtered_structure, left_on='ส่วนงาน', right_on='อักษรย่อ', how='left')

print(merged_data_structure.head())

# เปลี่ยนชื่อคอลัมน์ 'ต.บริหาร' เป็น 'ตำแหน่ง'
merged_data = merged_data.rename(columns={'ต.บริหาร': 'ตำแหน่ง'})
print(merged_data_structure.head())


# จัดเรียงลำดับคอลัมน์ใหม่
new_column_order = [
    'prefix', 'first_name', 'last_name', 'รหัสพนักงาน', 'ชื่อ-นามสกุล', 'ชื่อ-อังกฤษ', 'นามสกุล-อังกฤษ', 'ตำแหน่ง', 
    'ตำแหน่ง', 'ส่วนงาน', 'ชื่อเต็มส่วนงาน', 'ฝ่าย', 'กลุ่มงาน','สายงาน','e-mail', 'โทรศัพท์', 'มือถือ' 
]

# จัดเรียงคอลัมน์ใน DataFrame
reordered_data = merged_data_structure[new_column_order]

reordered_data.to_excel('update_staff_data.xlsx', index=False)
print(reordered_data.head())





