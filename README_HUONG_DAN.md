# Bot Điểm Danh Tuyết Nguyệt — Liên Đài và CTC

## Cấu hình Discord

- Server ID: `1528460047904936016`
- Voice Channel ID: `1528460048404320394`
- Kênh báo cáo: `1528460048404320390`
- Múi giờ: `Asia/Ho_Chi_Minh`

Bot tự vào room voice khi có ca đang chạy, tự mute/deafen và tự rời khi không còn ca nào. Bot chỉ tính thời gian khi một ca hoặc hoạt động điểm danh đang chạy.

## Ca cố định

| Ca | Giờ | Mức đạt |
|---|---:|---:|
| Boss Trưa | 12:00–12:40 | 15 phút |
| Vận Tiêu | 20:00–20:40 | 10 phút |
| Boss Tối | 22:00–22:40 | 15 phút |

## Ca Boss/Vận Tiêu linh động

Dùng trong phần chat của đúng room voice:

- `/vantieu` và `/ketthucvantieu`
- `/bosstrua` và `/ketthucbosstrua`
- `/bosstoi` và `/ketthucbosstoi`

Mức đạt bằng `3/5` tổng thời gian ca linh động.

Nếu mở ca linh động khi ca cố định cùng loại đang chạy, bot hủy hoàn toàn ca cố định, bỏ thời gian đã ghi nhận của ca cố định và chỉ chấm ca linh động. Kết quả mới thay thế toàn bộ kết quả cũ trong đúng cột của ngày đó.

## Liên Đài và CTC

Ba hoạt động này dùng **một lệnh để bắt đầu và dùng lại chính lệnh đó để kết thúc**:

| Hoạt động | Lệnh | Ngày | Cột Excel |
|---|---|---:|---:|
| Liên Đài Thứ 2 | `/liendaithu2` | Thứ 2 | H |
| Liên Đài Thứ 5 | `/liendaithu5` | Thứ 5 | R |
| CTC | `/ctc` | Thứ 7 | Y |

Cách dùng, ví dụ Liên Đài Thứ 2:

1. Quản lý đang ngồi trong đúng room voice và gõ `/liendaithu2` trong chat của room.
2. Bot vào voice và bắt đầu tính thời gian.
3. Khi hoạt động kết thúc, gõ lại `/liendaithu2`.
4. Người có mặt ít nhất `3/5` tổng thời gian hoạt động được ghi `1` vào cột H.

`/liendaithu5` và `/ctc` hoạt động tương tự. Lệnh chỉ cho bắt đầu vào đúng thứ; một hoạt động đã bắt đầu vẫn có thể kết thúc sau nửa đêm bằng chính lệnh đó.

Mỗi lần chốt hoạt động sẽ xóa kết quả cũ của đúng cột hoạt động đó rồi ghi lại kết quả mới.

## Lệnh `/tuanmoi`

Lệnh:

```text
/tuanmoi xac_nhan:True
```

Bot sẽ:

1. Sao lưu file tuần cũ vào `data/archives/`.
2. Xóa toàn bộ vùng điểm danh từ cột `E` đến `AB` của tất cả thành viên.
3. Vì vậy **Liên Đài Thứ 2, Liên Đài Thứ 5 và CTC cũng bị xóa**.
4. Giữ nguyên danh sách thành viên và công thức tổng ở cột `AC`, `AD`.
5. Xóa dữ liệu ca của tuần hiện tại trong SQLite.

## Form Excel

- File mẫu: `template/Tuyet_Nguyet.xlsx`
- File đang làm việc: `data/Tuyet_Nguyet.xlsx`
- Sheet: `Trang tính2`
- STT: cột `A`
- Thành viên bắt đầu từ hàng `3`
- Kết thúc danh sách tại dòng có chữ `Số Người Tham Gia Hoạt Động`

Các cột:

| Ngày | Boss Trưa | Vận Tiêu | Boss Tối | Hoạt động riêng |
|---|---|---|---|---|
| Thứ 2 | E | F | G | Liên Đài: H |
| Thứ 3 | I | J | K | — |
| Thứ 4 | L | M | N | — |
| Thứ 5 | O | P | Q | Liên Đài: R |
| Thứ 6 | S | T | U | — |
| Thứ 7 | V | W | X | CTC: Y |
| Chủ nhật | Z | AB | AA | — |

## Thêm thành viên

1. Mở `data/Tuyet_Nguyet.xlsx`.
2. Chèn hàng mới ngay phía trên dòng `Số Người Tham Gia Hoạt Động`.
3. Sao chép định dạng và công thức từ hàng thành viên phía trên.
4. Điền STT mới ở cột A và tên ở cột B.
5. Lưu rồi đóng Excel.

Bot tự quét toàn bộ STT dương đến dòng tổng kết nên không giới hạn số thành viên. STT không được trùng.

Tên Discord phải có STT ở đầu, ví dụ:

- `73 Tên thành viên`
- `73-Tên thành viên`
- `[73] Tên thành viên`

## Cài và chạy Windows

1. Chạy `1_CAI_DAT.bat`.
2. Điền token vào `.env`:

```env
DISCORD_BOT_TOKEN=token_thật
```

3. Chạy `2_KIEM_TRA.bat`.
4. Chạy `4_CHAY_BOT.bat`.

Bot cần các quyền: View Channel, Connect, Send Messages, Read Message History, Attach Files và Use Application Commands.

## Railway

- Không đưa `.env` lên GitHub.
- Railway Variable: `DISCORD_BOT_TOKEN`.
- Volume mount: `/app/data`.
- Sau khi cập nhật GitHub, Railway tự deploy lại.
