using System.Data;
using System.Globalization;
using ClosedXML.Excel;
using DntuLab.Api.Data;
using DntuLab.Api.Domain;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;

namespace DntuLab.Api.Controllers;

public record LichHocInput(string TenDangNhap, DateOnly Ngay, string Ca, TimeOnly BatDau, TimeOnly KetThuc, string MonHoc, string MaPhong);
public record ImportLichHocRequest(List<LichHocInput> Lich);
public record LichHocRow(int Dong, LichHocInput? DuLieu, List<string> Loi);

[ApiController]
[Route("api/quan-ly-lich-hoc")]
[Authorize(Roles = Roles.QL)]
public class QuanLyLichHocController(LabDbContext db) : ControllerBase
{
    [HttpGet("danh-muc")]
    public async Task<IActionResult> DanhMuc() => Ok(new {
        giangVien = await db.NguoiDungs.AsNoTracking().Where(x => x.VaiTro == Roles.GV && x.HoatDong)
            .Select(x => new { x.TenDangNhap, x.HoTen }).ToListAsync(),
        phong = await db.Phongs.AsNoTracking().Select(x => new { x.MaPhong, x.LoaiPhong }).ToListAsync()
    });

    [HttpGet("mau-excel")]
    public IActionResult MauExcel()
    {
        using var workbook = new XLWorkbook();
        var ws = workbook.AddWorksheet("Import_LichHoc");
        var headers = new[] { "STT", "TaiKhoanGiangVien", "NgayHoc", "CaHoc", "GioBatDau", "GioKetThuc", "MonHoc", "MaPhong", "GhiChu" };
        for (var i = 0; i < headers.Length; i++) ws.Cell(1, i + 1).Value = headers[i];
        ws.Range(1, 1, 1, 9).Style.Font.Bold = true;
        ws.Range(1, 1, 1, 9).Style.Fill.BackgroundColor = XLColor.DarkBlue;
        ws.Range(1, 1, 1, 9).Style.Font.FontColor = XLColor.White;
        ws.Cell(2, 1).Value = 1;
        ws.Cell(2, 2).Value = "gv01";
        ws.Cell(2, 3).Value = "2026-11-16";
        ws.Cell(2, 4).Value = "Ca 1";
        ws.Cell(2, 5).Value = "07:00";
        ws.Cell(2, 6).Value = "09:00";
        ws.Cell(2, 7).Value = "Lập trình C#";
        ws.Cell(2, 8).Value = "F507";
        ws.Columns().AdjustToContents();
        ws.SheetView.FreezeRows(1);
        var guide = workbook.AddWorksheet("HuongDan");
        guide.Cell(1, 1).Value = "Điền dữ liệu từ dòng 2; xóa dòng ví dụ nếu không muốn nhập.";
        guide.Cell(2, 1).Value = "Ngày: yyyy-MM-dd hoặc dd/MM/yyyy. Giờ: HH:mm.";
        guide.Cell(3, 1).Value = "Tên đăng nhập phải thuộc giảng viên đang hoạt động; mã phòng phải tồn tại.";
        guide.Column(1).Width = 95;
        using var ms = new MemoryStream();
        workbook.SaveAs(ms);
        return File(ms.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "DNTU_Lab_Mau_Import_LichHoc.xlsx");
    }

    [HttpPost("xem-truoc")]
    [RequestSizeLimit(5_000_000)]
    public async Task<IActionResult> XemTruoc(IFormFile file)
    {
        if (file == null || file.Length == 0 || file.Length > 5_000_000 ||
            !string.Equals(Path.GetExtension(file.FileName), ".xlsx", StringComparison.OrdinalIgnoreCase))
            return BadRequest(new { message = "Chỉ chấp nhận file .xlsx tối đa 5MB." });
        try {
            using var stream = file.OpenReadStream();
            using var workbook = new XLWorkbook(stream);
            if (!workbook.Worksheets.TryGetWorksheet("Import_LichHoc", out var ws))
                return BadRequest(new { message = "Thiếu sheet Import_LichHoc." });
            var last = ws.LastRowUsed()?.RowNumber() ?? 1;
            if (last > 1001) return BadRequest(new { message = "Tối đa 1000 dòng dữ liệu." });
            var result = new List<LichHocRow>();
            for (int row = 2; row <= last; row++) {
                if (Enumerable.Range(2, 7).All(col => string.IsNullOrWhiteSpace(ws.Cell(row, col).GetFormattedString()))) continue;
                var errors = new List<string>();
                string V(int col) => ws.Cell(row, col).GetFormattedString().Trim();
                var dateText = V(3);
                var timeStart = V(5);
                var timeEnd = V(6);
                var validDate = DateOnly.TryParseExact(dateText, new[] { "yyyy-MM-dd", "dd/MM/yyyy", "d/M/yyyy" }, CultureInfo.InvariantCulture, DateTimeStyles.None, out var date);
                if (!validDate && ws.Cell(row, 3).TryGetValue<DateTime>(out var excelDate)) { date = DateOnly.FromDateTime(excelDate); validDate = true; }
                var validStart = TimeOnly.TryParse(timeStart, CultureInfo.InvariantCulture, DateTimeStyles.None, out var start);
                var validEnd = TimeOnly.TryParse(timeEnd, CultureInfo.InvariantCulture, DateTimeStyles.None, out var end);
                if (!validDate) errors.Add("Ngày học không hợp lệ");
                if (!validStart || !validEnd) errors.Add("Giờ học không hợp lệ");
                var data = validDate && validStart && validEnd
                    ? new LichHocInput(V(2), date, V(4), start, end, V(7), V(8)) : null;
                result.Add(new LichHocRow(row, data, errors));
            }
            var checkedRows = await Validate(result);
            return Ok(new { rows = checkedRows, hopLe = checkedRows.Count(x => x.Loi.Count == 0), loi = checkedRows.Count(x => x.Loi.Count > 0) });
        } catch (Exception) {
            return BadRequest(new { message = "Không đọc được file Excel. Hãy dùng file mẫu của hệ thống." });
        }
    }

    [HttpPost("tao")]
    public async Task<IActionResult> Tao([FromBody] LichHocInput input) => await Save(new List<LichHocInput> { input });

    [HttpPost("import")]
    public async Task<IActionResult> Import([FromBody] ImportLichHocRequest request) => await Save(request.Lich);

    private async Task<IActionResult> Save(List<LichHocInput>? inputs)
    {
        if (inputs == null || inputs.Count == 0 || inputs.Count > 1000)
            return BadRequest(new { message = "Số lịch phải từ 1 đến 1000." });
        await using var tx = await db.Database.BeginTransactionAsync(IsolationLevel.Serializable);
        var checkedRows = await Validate(inputs.Select((x, i) => new LichHocRow(i + 2, x, new())).ToList());
        if (checkedRows.Any(x => x.Loi.Count > 0))
            return Conflict(new { message = "Có lịch không hợp lệ hoặc xung đột. Không có lịch nào được lưu.", rows = checkedRows });
        var accounts = await db.NguoiDungs.Where(x => x.VaiTro == Roles.GV && x.HoatDong).ToDictionaryAsync(x => x.TenDangNhap);
        foreach (var x in inputs) db.LichDays.Add(new LichDay {
            GiangVienId = accounts[x.TenDangNhap.Trim()].Id,
            Ngay = x.Ngay, Ca = x.Ca.Trim(), BatDau = x.BatDau, KetThuc = x.KetThuc,
            MonHoc = x.MonHoc.Trim(), MaPhong = x.MaPhong.Trim(),
            TrangThai = TrangThaiLich.ChuaCheckIn, SoBuoi = 1
        });
        await db.SaveChangesAsync();
        await tx.CommitAsync();
        return Ok(new { success = true, daTao = inputs.Count });
    }

    private async Task<List<LichHocRow>> Validate(List<LichHocRow> rows)
    {
        var accounts = await db.NguoiDungs.AsNoTracking().Where(x => x.VaiTro == Roles.GV && x.HoatDong)
            .Select(x => new { x.Id, x.TenDangNhap }).ToListAsync();
        var rooms = await db.Phongs.AsNoTracking().Select(x => x.MaPhong).ToListAsync();
        var accountMap = accounts.ToDictionary(x => x.TenDangNhap, x => x.Id, StringComparer.OrdinalIgnoreCase);
        var roomSet = rooms.ToHashSet(StringComparer.OrdinalIgnoreCase);
        var validDates = rows.Where(x => x.DuLieu != null).Select(x => x.DuLieu!.Ngay).Distinct().ToList();
        var existing = await db.LichDays.AsNoTracking().Where(x => validDates.Contains(x.Ngay))
            .Select(x => new { x.Ngay, x.BatDau, x.KetThuc, x.MaPhong, x.GiangVienId }).ToListAsync();
        var result = new List<LichHocRow>();
        foreach (var row in rows) {
            var errors = new List<string>(row.Loi);
            var x = row.DuLieu;
            if (x != null) {
                var user = (x.TenDangNhap ?? "").Trim();
                var room = (x.MaPhong ?? "").Trim();
                if (!accountMap.TryGetValue(user, out var teacherId)) errors.Add("Giảng viên không tồn tại/không hoạt động");
                if (!roomSet.Contains(room)) errors.Add("Phòng không tồn tại");
                if (string.IsNullOrWhiteSpace(x.MonHoc) || x.MonHoc.Length > 150) errors.Add("Môn học bắt buộc, tối đa 150 ký tự");
                if (string.IsNullOrWhiteSpace(x.Ca) || x.Ca.Length > 10) errors.Add("Ca học bắt buộc, tối đa 10 ký tự");
                if (x.BatDau >= x.KetThuc) errors.Add("Giờ bắt đầu phải trước giờ kết thúc");
                if (x.Ngay < DateOnly.FromDateTime(DateTime.Today)) errors.Add("Không tạo lịch trong quá khứ");
                bool Overlap(DateOnly d, TimeOnly s, TimeOnly e) => x.Ngay == d && x.BatDau < e && s < x.KetThuc;
                if (existing.Any(y => Overlap(y.Ngay, y.BatDau, y.KetThuc) &&
                    (string.Equals(y.MaPhong, room, StringComparison.OrdinalIgnoreCase) || y.GiangVienId == teacherId)))
                    errors.Add("Trùng lịch phòng hoặc giảng viên trong database");
                if (rows.Any(y => y.Dong != row.Dong && y.DuLieu is { } other &&
                    Overlap(other.Ngay, other.BatDau, other.KetThuc) &&
                    (string.Equals(other.MaPhong?.Trim(), room, StringComparison.OrdinalIgnoreCase) ||
                     string.Equals(other.TenDangNhap?.Trim(), user, StringComparison.OrdinalIgnoreCase))))
                    errors.Add("Trùng phòng hoặc giảng viên trong danh sách import");
            }
            result.Add(new LichHocRow(row.Dong, x, errors));
        }
        return result;
    }
}
