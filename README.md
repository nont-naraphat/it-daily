# IT Daily

หน้า dashboard ส่วนตัวที่ผูกกับ Microsoft Planner แล้วให้ AI สรุปเคสรายวันให้เองทุกเช้า
(FastAPI + vanilla JS + SQLite, คอนเทนเนอร์เดียว)

## 1. สร้าง App Registration ใน Entra
Entra admin center → App registrations → New registration
- Name: `IT Daily`
- Supported account types: Single tenant
- ไม่ต้องใส่ Redirect URI

หลังสร้าง:
- **Authentication** → เปิด `Allow public client flows` = **Yes** (จำเป็นสำหรับ device-code flow)
- **API permissions** (Delegated) → เพิ่ม `Tasks.ReadWrite`, `Group.Read.All`, `User.Read` → **Grant admin consent** (ขอโมกด)
- คัดลอก **Application (client) ID** ไปใส่ `.env`

> ไม่ต้องมี client secret — auth เป็นตัวพี่เองผ่าน device code แล้ว refresh token ถูก cache ไว้ใน `data/token_cache.json` ให้ scheduler ใช้เองข้ามวัน

## 2. ตั้งค่า
```
cp .env.example .env
```
แก้ `CLIENT_ID` และ `ANTHROPIC_API_KEY` (TENANT_ID ใส่ให้แล้ว)

## 3. Deploy
```
docker compose up -d --build
```
เปิด `http://<synology-ip>:8094`

## 4. เชื่อมครั้งแรก
กด **Connect** → เอารหัสไปกรอกที่ `microsoft.com/devicelogin` → เสร็จแล้วหน้าเว็บจะดึงงานเอง
ทุกวันจันทร์–ศุกร์ เวลา `BRIEF_HOUR:BRIEF_MINUTE` scheduler จะดึง Planner + สรุปใหม่ให้อัตโนมัติ
(กด **Sync now** เพื่อรีเฟรชทันทีได้)

## ทำอะไรได้
- ดึง task ที่ถูก assign ให้บัญชีที่ connect (`/me/planner/tasks`) → เห็นทั้ง plan เจ้านายและ IT workstream ที่พี่มีส่วนร่วม
- **สร้าง task ใหม่** จากในแอป → เลือก plan/bucket, กำหนดส่ง, assign ให้ตัวเอง → เขียนเข้า Planner จริง
- **อัปเดตความคืบหน้า / ปิดงาน** (0/25/50/75/100%) จากการ์ดแต่ละใบ → sync กลับ Planner ทันที
- AI สรุปเคสรายวันให้เองทุกเช้า

> การสร้าง/อัปเดตมีผลกับ Planner จริงทันที (ไม่ใช่ draft) — งานที่สร้าง/แก้จะไปโผล่ใน Planner ของทุกคนที่เห็น plan นั้นด้วย
> ใช้ delegated `Tasks.ReadWrite` ที่ขอไว้แล้ว ไม่ต้องเพิ่ม permission
> ถ้าสร้างงานใน plan แล้วเจอ 403 = บัญชีที่ connect ต้องเป็นสมาชิกของ group ที่เป็นเจ้าของ plan นั้น
- ถ้า tenant บังคับ sign-in frequency กับบัญชี admin อาจต้อง Connect ซ้ำเป็นระยะ → แนะนำ auth ด้วยบัญชีที่ไม่โดน CA บีบ หรือ exclude แอปนี้
- token cache + SQLite เก็บใน `./data` (mount ไว้แล้วใน compose)
