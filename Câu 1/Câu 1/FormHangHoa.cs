using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Threading;
using System.Windows.Forms;
using ClosedXML.Excel;
using FastReport;
using FastReport.Export.Pdf;
using Câu_1.Models;

namespace Câu_1
{
    public partial class FormHangHoa : Form
    {
        private TextBox txtTen = new TextBox();
        private ComboBox cboLoai = new ComboBox();
        private ComboBox cboDonVi = new ComboBox();
        private NumericUpDown numGia = new NumericUpDown();
        private NumericUpDown numTon = new NumericUpDown();
        private DateTimePicker dtNgay = new DateTimePicker();
        private CheckBox chkKinhDoanh = new CheckBox();

        private TextBox txtTim = new TextBox();
        private ComboBox cboLoc = new ComboBox();
        private CheckBox chkConHang = new CheckBox();

        private DataGridView dgv = new DataGridView();
        private Label lblThongBao = new Label();

        private System.Windows.Forms.Timer searchTimer = new System.Windows.Forms.Timer();
        private bool isBinding = false;
        private int? selectedMaHang = null;

        public FormHangHoa()
        {
            InitializeComponent();

            // Cài đặt định dạng tiếng Việt cho tiền tệ và ngày tháng
            CultureInfo viCulture = new CultureInfo("vi-VN");
            Thread.CurrentThread.CurrentCulture = viCulture;
            Thread.CurrentThread.CurrentUICulture = viCulture;

            Text = "Quản lý hàng hóa - Tạp hóa UTC2";
            StartPosition = FormStartPosition.CenterScreen;
            ClientSize = new Size(830, 465);
            MinimumSize = new Size(845, 505);
            Font = new Font("Segoe UI", 9F);
            BackColor = Color.WhiteSmoke;

            TaoGiaoDien();

            // Khởi tạo Timer tìm kiếm (350 ms debounce)
            searchTimer.Interval = 350;
            searchTimer.Tick += (s, e) =>
            {
                searchTimer.Stop();
                LoadDuLieu();
            };

            Load += (s, e) =>
            {
                LoadDuLieu();
                LamMoi();
            };
        }

        private void TaoGiaoDien()
        {
            // --- GroupBox Thông tin hàng hóa ---
            GroupBox gb = new GroupBox
            {
                Text = "Thông tin hàng hóa",
                Location = new Point(12, 10),
                Size = new Size(490, 140)
            };
            Controls.Add(gb);

            ThemLabel(gb, "Tên hàng:", 12, 22);
            txtTen.SetBounds(82, 20, 395, 23);
            gb.Controls.Add(txtTen);

            ThemLabel(gb, "Loại hàng:", 12, 50);
            cboLoai.SetBounds(82, 47, 150, 23);
            cboLoai.DropDownStyle = ComboBoxStyle.DropDownList;
            cboLoai.Items.AddRange(new object[] { "Đồ uống", "Thực phẩm khô", "Gia vị", "Đồ gia dụng" });
            gb.Controls.Add(cboLoai);

            ThemLabel(gb, "Số lượng tồn:", 248, 50);
            numTon.SetBounds(342, 47, 135, 23);
            numTon.Maximum = 1000000;
            numTon.Minimum = 0;
            gb.Controls.Add(numTon);

            ThemLabel(gb, "Đơn vị:", 12, 78);
            cboDonVi.SetBounds(82, 75, 150, 23);
            cboDonVi.DropDownStyle = ComboBoxStyle.DropDownList;
            cboDonVi.Items.AddRange(new object[] { "Chai", "Gói", "Hộp", "Túi" });
            gb.Controls.Add(cboDonVi);

            ThemLabel(gb, "Ngày nhập:", 248, 78);
            dtNgay.SetBounds(342, 75, 135, 23);
            dtNgay.Format = DateTimePickerFormat.Custom;
            dtNgay.CustomFormat = "dd/MM/yyyy";
            gb.Controls.Add(dtNgay);

            ThemLabel(gb, "Giá bán:", 12, 106);
            numGia.SetBounds(82, 103, 150, 23);
            numGia.Maximum = 1000000000;
            numGia.Increment = 500;
            numGia.ThousandsSeparator = true;
            gb.Controls.Add(numGia);

            chkKinhDoanh.Text = "Còn kinh doanh";
            chkKinhDoanh.SetBounds(342, 104, 135, 23);
            chkKinhDoanh.Checked = true;
            gb.Controls.Add(chkKinhDoanh);

            // --- GroupBox Tìm kiếm ---
            GroupBox gbTim = new GroupBox
            {
                Text = "Tìm kiếm",
                Location = new Point(510, 10),
                Size = new Size(308, 140)
            };
            Controls.Add(gbTim);

            txtTim.SetBounds(12, 20, 282, 23);
            txtTim.PlaceholderText = "Nhập tên hàng cần tìm...";
            txtTim.TextChanged += (s, e) =>
            {
                searchTimer.Stop();
                searchTimer.Start();
            };
            gbTim.Controls.Add(txtTim);

            cboLoc.SetBounds(12, 47, 282, 23);
            cboLoc.DropDownStyle = ComboBoxStyle.DropDownList;
            cboLoc.Items.AddRange(new object[] { "Tất cả loại", "Đồ uống", "Thực phẩm khô", "Gia vị", "Đồ gia dụng" });
            cboLoc.SelectedIndex = 0;
            cboLoc.SelectedIndexChanged += (s, e) => LoadDuLieu();
            gbTim.Controls.Add(cboLoc);

            chkConHang.Text = "Chỉ hiện hàng còn kinh doanh";
            chkConHang.SetBounds(12, 77, 282, 25);
            chkConHang.CheckedChanged += (s, e) => LoadDuLieu();
            gbTim.Controls.Add(chkConHang);

            // --- DataGridView ---
            dgv.SetBounds(12, 158, 806, 235);
            dgv.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
            dgv.BackgroundColor = Color.DarkGray;
            dgv.AllowUserToAddRows = false;
            dgv.AllowUserToDeleteRows = false;
            dgv.ReadOnly = true;
            dgv.MultiSelect = false;
            dgv.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            dgv.AutoGenerateColumns = false;
            dgv.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            dgv.RowHeadersWidth = 38;

            ThemCot("MaHang", "Mã", 45, "", DataGridViewContentAlignment.MiddleCenter);
            ThemCot("TenHang", "Tên hàng", 160, "", DataGridViewContentAlignment.MiddleLeft);
            ThemCot("LoaiHang", "Loại hàng", 85, "", DataGridViewContentAlignment.MiddleLeft);
            ThemCot("DonVi", "Đơn vị", 55, "", DataGridViewContentAlignment.MiddleCenter);
            ThemCot("GiaBan", "Giá bán", 80, "N0", DataGridViewContentAlignment.MiddleRight);
            ThemCot("SoLuongTon", "Tồn", 55, "N0", DataGridViewContentAlignment.MiddleRight);
            ThemCot("NgayNhap", "Ngày nhập", 85, "dd/MM/yyyy", DataGridViewContentAlignment.MiddleCenter);

            dgv.Columns.Add(new DataGridViewCheckBoxColumn
            {
                DataPropertyName = "ConKinhDoanh",
                HeaderText = "Kinh doanh",
                FillWeight = 75,
                DefaultCellStyle = { Alignment = DataGridViewContentAlignment.MiddleCenter }
            });

            dgv.SelectionChanged += (s, e) => DoLenForm();
            dgv.CellClick += (s, e) => DoLenForm();
            Controls.Add(dgv);

            // --- Các nút chức năng ---
            int btnY = 400;
            TaoNut("Thêm", 12, btnY, 76, Them);
            TaoNut("Sửa", 94, btnY, 76, Sua);
            TaoNut("Xóa", 176, btnY, 76, Xoa);
            TaoNut("Làm mới", 258, btnY, 80, LamMoi);
            TaoNut("Xem bảng giá", 344, btnY, 95, XemBangGia);
            TaoNut("Xuất PDF", 445, btnY, 82, XuatPDF);
            TaoNut("Xuất Excel (lưới)", 533, btnY, 130, XuatExcelLuoi);
            TaoNut("Xuất Excel tồn kho", 669, btnY, 145, XuatExcelTonKho);

            // --- Label thông báo trạng thái ---
            lblThongBao.SetBounds(12, 435, 806, 22);
            lblThongBao.Anchor = AnchorStyles.Bottom | AnchorStyles.Left;
            Controls.Add(lblThongBao);
        }

        private void ThemLabel(Control cha, string text, int x, int y)
        {
            cha.Controls.Add(new Label
            {
                Text = text,
                Location = new Point(x, y),
                AutoSize = true
            });
        }

        private void ThemCot(string property, string title, int width, string format, DataGridViewContentAlignment align)
        {
            dgv.Columns.Add(new DataGridViewTextBoxColumn
            {
                DataPropertyName = property,
                HeaderText = title,
                FillWeight = width,
                DefaultCellStyle =
                {
                    Format = format,
                    Alignment = align
                }
            });
        }

        private void TaoNut(string text, int x, int y, int width, Action action)
        {
            Button btn = new Button
            {
                Text = text,
                Location = new Point(x, y),
                Size = new Size(width, 28),
                Anchor = AnchorStyles.Bottom | AnchorStyles.Left
            };
            btn.Click += (s, e) => action();
            Controls.Add(btn);
        }

        private void LoadDuLieu()
        {
            try
            {
                using var db = new TiemTapHoaDbContext();
                var query = db.HangHoas.AsQueryable();

                string tuKhoa = txtTim.Text.Trim();
                if (!string.IsNullOrEmpty(tuKhoa))
                {
                    query = query.Where(h => h.TenHang.Contains(tuKhoa));
                }

                if (cboLoc.SelectedIndex > 0)
                {
                    string loai = cboLoc.SelectedItem?.ToString() ?? "";
                    query = query.Where(h => h.LoaiHang == loai);
                }

                if (chkConHang.Checked)
                {
                    query = query.Where(h => h.ConKinhDoanh);
                }

                var list = query.OrderBy(h => h.MaHang).ToList();

                isBinding = true;
                dgv.DataSource = null;
                dgv.DataSource = list;
                isBinding = false;

                lblThongBao.Text = $"Tìm thấy {list.Count} mặt hàng";

                if (selectedMaHang.HasValue)
                {
                    bool found = false;
                    foreach (DataGridViewRow r in dgv.Rows)
                    {
                        if (r.DataBoundItem is HangHoa h && h.MaHang == selectedMaHang.Value)
                        {
                            r.Selected = true;
                            dgv.CurrentCell = r.Cells[0];
                            found = true;
                            break;
                        }
                    }
                    if (!found)
                    {
                        dgv.ClearSelection();
                    }
                }
                else
                {
                    dgv.ClearSelection();
                }
            }
            catch (Exception ex)
            {
                lblThongBao.Text = "Lỗi tải dữ liệu: " + ex.Message;
            }
        }

        private HangHoa? DangChon()
        {
            return dgv.CurrentRow?.DataBoundItem as HangHoa;
        }

        private void DoLenForm()
        {
            if (isBinding) return;
            HangHoa? h = DangChon();
            if (h == null) return;

            selectedMaHang = h.MaHang;
            txtTen.Text = h.TenHang ?? "";

            cboLoai.SelectedItem = h.LoaiHang;
            if (cboLoai.SelectedIndex < 0 && !string.IsNullOrEmpty(h.LoaiHang))
                cboLoai.Text = h.LoaiHang;

            cboDonVi.SelectedItem = h.DonVi;
            if (cboDonVi.SelectedIndex < 0 && !string.IsNullOrEmpty(h.DonVi))
                cboDonVi.Text = h.DonVi;

            numGia.Value = Math.Min(numGia.Maximum, Math.Max(numGia.Minimum, h.GiaBan ?? 0));
            numTon.Value = Math.Min(numTon.Maximum, Math.Max(numTon.Minimum, h.SoLuongTon ?? 0));
            dtNgay.Value = h.NgayNhap ?? DateTime.Today;
            chkKinhDoanh.Checked = h.ConKinhDoanh;
        }

        private bool KiemTra()
        {
            if (string.IsNullOrWhiteSpace(txtTen.Text))
            {
                MessageBox.Show("Vui lòng nhập tên hàng!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                txtTen.Focus();
                return false;
            }

            if (numGia.Value <= 0)
            {
                MessageBox.Show("Giá bán phải lớn hơn 0!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                numGia.Focus();
                return false;
            }

            if (cboLoai.SelectedIndex < 0 && string.IsNullOrWhiteSpace(cboLoai.Text))
            {
                MessageBox.Show("Vui lòng chọn loại hàng!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                cboLoai.Focus();
                return false;
            }

            if (cboDonVi.SelectedIndex < 0 && string.IsNullOrWhiteSpace(cboDonVi.Text))
            {
                MessageBox.Show("Vui lòng chọn đơn vị tính!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                cboDonVi.Focus();
                return false;
            }

            return true;
        }

        private void Them()
        {
            if (!KiemTra()) return;

            try
            {
                using var db = new TiemTapHoaDbContext();
                var hh = new HangHoa
                {
                    TenHang = txtTen.Text.Trim(),
                    LoaiHang = cboLoai.Text.Trim(),
                    DonVi = cboDonVi.Text.Trim(),
                    GiaBan = numGia.Value,
                    SoLuongTon = (int)numTon.Value,
                    NgayNhap = dtNgay.Value.Date,
                    ConKinhDoanh = chkKinhDoanh.Checked
                };

                db.HangHoas.Add(hh);
                db.SaveChanges();

                selectedMaHang = hh.MaHang;
                MessageBox.Show("Thêm hàng hóa thành công!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                LoadDuLieu();
                LamMoi();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi khi thêm hàng hóa: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void Sua()
        {
            if (selectedMaHang == null)
            {
                MessageBox.Show("Vui lòng chọn hàng cần sửa!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (!KiemTra()) return;

            try
            {
                using var db = new TiemTapHoaDbContext();
                var hh = db.HangHoas.Find(selectedMaHang.Value);
                if (hh == null)
                {
                    MessageBox.Show("Không tìm thấy hàng hóa cần sửa trong CSDL!", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                hh.TenHang = txtTen.Text.Trim();
                hh.LoaiHang = cboLoai.Text.Trim();
                hh.DonVi = cboDonVi.Text.Trim();
                hh.GiaBan = numGia.Value;
                hh.SoLuongTon = (int)numTon.Value;
                hh.NgayNhap = dtNgay.Value.Date;
                hh.ConKinhDoanh = chkKinhDoanh.Checked;

                db.SaveChanges();

                MessageBox.Show("Cập nhật thành công!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                LoadDuLieu();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi khi cập nhật hàng hóa: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void Xoa()
        {
            if (selectedMaHang == null)
            {
                MessageBox.Show("Vui lòng chọn hàng cần xóa!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            string tenHang = txtTen.Text.Trim();
            if (MessageBox.Show($"Bạn có chắc muốn xóa {tenHang}?", "Xác nhận", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                try
                {
                    using var db = new TiemTapHoaDbContext();
                    var hh = db.HangHoas.Find(selectedMaHang.Value);
                    if (hh != null)
                    {
                        db.HangHoas.Remove(hh);
                        db.SaveChanges();
                        MessageBox.Show("Xóa hàng hóa thành công!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        selectedMaHang = null;
                        LoadDuLieu();
                        LamMoi();
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Lỗi khi xóa hàng hóa: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private void LamMoi()
        {
            selectedMaHang = null;
            txtTen.Clear();
            cboLoai.SelectedIndex = -1;
            cboLoai.Text = "";
            cboDonVi.SelectedIndex = -1;
            cboDonVi.Text = "";
            numGia.Value = 0;
            numTon.Value = 0;
            dtNgay.Value = DateTime.Today;
            chkKinhDoanh.Checked = true;
            dgv.ClearSelection();
            txtTen.Focus();
        }

        private string GetReportPath()
        {
            string p1 = Path.Combine(Application.StartupPath, "Reports", "BangGia.frx");
            if (File.Exists(p1)) return p1;
            string p2 = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, @"..\..\..\Reports\BangGia.frx");
            if (File.Exists(p2)) return Path.GetFullPath(p2);
            return p1;
        }

        private DataTable TaoDataTableHangHoa(List<HangHoa> list)
        {
            var dt = new DataTable("HangHoa");
            dt.Columns.Add("MaHang", typeof(int));
            dt.Columns.Add("TenHang", typeof(string));
            dt.Columns.Add("LoaiHang", typeof(string));
            dt.Columns.Add("DonVi", typeof(string));
            dt.Columns.Add("GiaBan", typeof(decimal));
            dt.Columns.Add("SoLuongTon", typeof(int));
            dt.Columns.Add("NgayNhap", typeof(DateTime));
            dt.Columns.Add("ConKinhDoanh", typeof(bool));

            foreach (var h in list)
            {
                dt.Rows.Add(
                    h.MaHang,
                    h.TenHang,
                    h.LoaiHang,
                    h.DonVi,
                    h.GiaBan ?? 0,
                    h.SoLuongTon ?? 0,
                    h.NgayNhap ?? DateTime.Today,
                    h.ConKinhDoanh
                );
            }
            return dt;
        }

        private void XemBangGia()
        {
            try
            {
                using var db = new TiemTapHoaDbContext();
                var dsConKinhDoanh = db.HangHoas
                    .Where(h => h.ConKinhDoanh)
                    .OrderBy(h => h.MaHang)
                    .ToList();

                var dt = TaoDataTableHangHoa(dsConKinhDoanh);

                string reportPath = GetReportPath();
                if (!File.Exists(reportPath))
                {
                    MessageBox.Show("Không tìm thấy file mẫu báo cáo BangGia.frx tại: " + reportPath, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                using var report = new Report();
                report.Load(reportPath);
                report.RegisterData(dt, "HangHoa");
                report.GetDataSource("HangHoa").Enabled = true;
                report.SetParameterValue("NgayIn", DateTime.Now.ToString("dd/MM/yyyy"));
                report.Show();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi khi mở bảng giá: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void XuatPDF()
        {
            try
            {
                using var db = new TiemTapHoaDbContext();
                var dsConKinhDoanh = db.HangHoas
                    .Where(h => h.ConKinhDoanh)
                    .OrderBy(h => h.MaHang)
                    .ToList();

                var dt = TaoDataTableHangHoa(dsConKinhDoanh);

                string reportPath = GetReportPath();
                if (!File.Exists(reportPath))
                {
                    MessageBox.Show("Không tìm thấy file mẫu báo cáo BangGia.frx!", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                using SaveFileDialog sfd = new SaveFileDialog
                {
                    Filter = "PDF (*.pdf)|*.pdf",
                    FileName = "BangGiaHangHoa.pdf"
                };

                if (sfd.ShowDialog() != DialogResult.OK) return;

                using var report = new Report();
                report.Load(reportPath);
                report.RegisterData(dt, "HangHoa");
                report.GetDataSource("HangHoa").Enabled = true;
                report.SetParameterValue("NgayIn", DateTime.Now.ToString("dd/MM/yyyy"));
                report.Prepare();

                var pdfExport = new PDFExport();
                report.Export(pdfExport, sfd.FileName);

                if (MessageBox.Show("Xuất file PDF thành công! Bạn có muốn mở file không?", "Xác nhận", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                {
                    Process.Start(new ProcessStartInfo(sfd.FileName) { UseShellExecute = true });
                }
            }
            catch (IOException)
            {
                MessageBox.Show("File PDF đang được mở bởi ứng dụng khác. Vui lòng đóng file và thử lại!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi khi xuất PDF: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void XuatExcelLuoi()
        {
            if (dgv.Rows.Count == 0)
            {
                MessageBox.Show("Không có dữ liệu trên lưới để xuất!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            using SaveFileDialog sfd = new SaveFileDialog
            {
                Filter = "Excel Workbook (*.xlsx)|*.xlsx",
                FileName = "DanhSachHangHoa.xlsx"
            };

            if (sfd.ShowDialog() != DialogResult.OK) return;

            try
            {
                using var workbook = new XLWorkbook();
                var ws = workbook.Worksheets.Add("HangHoa");

                // Tiêu đề các cột
                string[] headers = new string[] { "Mã", "Tên hàng", "Loại hàng", "Đơn vị", "Giá bán", "Tồn", "Ngày nhập", "Kinh doanh" };
                for (int c = 0; c < headers.Length; c++)
                {
                    var cell = ws.Cell(1, c + 1);
                    cell.Value = headers[c];
                    cell.Style.Font.Bold = true;
                    cell.Style.Fill.BackgroundColor = XLColor.LightGray;
                    cell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                }

                int rowIndex = 2;
                foreach (DataGridViewRow row in dgv.Rows)
                {
                    if (row.IsNewRow) continue;
                    if (row.DataBoundItem is HangHoa h)
                    {
                        // Số nguyên
                        ws.Cell(rowIndex, 1).Value = h.MaHang;
                        ws.Cell(rowIndex, 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

                        // Chuỗi
                        ws.Cell(rowIndex, 2).Value = h.TenHang ?? "";
                        ws.Cell(rowIndex, 3).Value = h.LoaiHang ?? "";
                        ws.Cell(rowIndex, 4).Value = h.DonVi ?? "";
                        ws.Cell(rowIndex, 4).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

                        // Giá bán (số thực/thập phân, căn phải, định dạng N0)
                        var cellGia = ws.Cell(rowIndex, 5);
                        cellGia.Value = (double)(h.GiaBan ?? 0);
                        cellGia.Style.NumberFormat.Format = "#,##0";
                        cellGia.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;

                        // Tồn kho (số nguyên, căn phải)
                        var cellTon = ws.Cell(rowIndex, 6);
                        cellTon.Value = h.SoLuongTon ?? 0;
                        cellTon.Style.NumberFormat.Format = "#,##0";
                        cellTon.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;

                        // Ngày nhập (DateTime)
                        var cellNgay = ws.Cell(rowIndex, 7);
                        if (h.NgayNhap.HasValue)
                        {
                            cellNgay.Value = h.NgayNhap.Value;
                            cellNgay.Style.DateFormat.Format = "dd/MM/yyyy";
                        }
                        cellNgay.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

                        // Kinh doanh
                        var cellKd = ws.Cell(rowIndex, 8);
                        cellKd.Value = h.ConKinhDoanh ? "Còn kinh doanh" : "Ngừng bán";
                        cellKd.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

                        rowIndex++;
                    }
                }

                var range = ws.Range(1, 1, rowIndex - 1, headers.Length);
                range.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                range.Style.Border.InsideBorder = XLBorderStyleValues.Thin;
                ws.Columns().AdjustToContents();

                workbook.SaveAs(sfd.FileName);

                if (MessageBox.Show("Xuất file Excel từ lưới thành công! Bạn có muốn mở file không?", "Xác nhận", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                {
                    Process.Start(new ProcessStartInfo(sfd.FileName) { UseShellExecute = true });
                }
            }
            catch (IOException)
            {
                MessageBox.Show("File Excel đang được mở bởi ứng dụng khác. Vui lòng đóng file và thử lại!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi khi xuất Excel: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void XuatExcelTonKho()
        {
            using SaveFileDialog sfd = new SaveFileDialog
            {
                Filter = "Excel Workbook (*.xlsx)|*.xlsx",
                FileName = "BaoCaoTonKho.xlsx"
            };

            if (sfd.ShowDialog() != DialogResult.OK) return;

            try
            {
                using var db = new TiemTapHoaDbContext();
                var dsConKinhDoanh = db.HangHoas
                    .Where(h => h.ConKinhDoanh)
                    .OrderBy(h => h.MaHang)
                    .ToList();

                using var workbook = new XLWorkbook();
                var ws = workbook.Worksheets.Add("TonKho");

                // Headers: STT, Mã hàng, Tên hàng, Loại hàng, Đơn vị, Giá bán, Số lượng tồn, Giá trị tồn
                string[] headers = new string[] { "STT", "Mã hàng", "Tên hàng", "Loại hàng", "Đơn vị", "Giá bán", "Số lượng tồn", "Giá trị tồn" };
                for (int c = 0; c < headers.Length; c++)
                {
                    var cell = ws.Cell(1, c + 1);
                    cell.Value = headers[c];
                    cell.Style.Font.Bold = true;
                    cell.Style.Fill.BackgroundColor = XLColor.LightGray;
                    cell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                }

                int rowIndex = 2;
                int stt = 1;
                foreach (var h in dsConKinhDoanh)
                {
                    // STT
                    ws.Cell(rowIndex, 1).Value = stt++;
                    ws.Cell(rowIndex, 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

                    // Mã hàng
                    ws.Cell(rowIndex, 2).Value = h.MaHang;
                    ws.Cell(rowIndex, 2).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

                    // Tên hàng
                    ws.Cell(rowIndex, 3).Value = h.TenHang ?? "";

                    // Loại hàng
                    ws.Cell(rowIndex, 4).Value = h.LoaiHang ?? "";

                    // Đơn vị
                    ws.Cell(rowIndex, 5).Value = h.DonVi ?? "";
                    ws.Cell(rowIndex, 5).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

                    // Giá bán (Cột F)
                    var cellGia = ws.Cell(rowIndex, 6);
                    cellGia.Value = (double)(h.GiaBan ?? 0);
                    cellGia.Style.NumberFormat.Format = "#,##0";
                    cellGia.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;

                    // Số lượng tồn (Cột G)
                    var cellTon = ws.Cell(rowIndex, 7);
                    cellTon.Value = h.SoLuongTon ?? 0;
                    cellTon.Style.NumberFormat.Format = "#,##0";
                    cellTon.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;

                    // Giá trị tồn (Cột H) = Giá bán * Số lượng tồn (Công thức)
                    var cellGiaTri = ws.Cell(rowIndex, 8);
                    cellGiaTri.FormulaA1 = $"=F{rowIndex}*G{rowIndex}";
                    cellGiaTri.Style.NumberFormat.Format = "#,##0";
                    cellGiaTri.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;

                    rowIndex++;
                }

                // Dòng Tổng
                int rowTong = rowIndex;
                ws.Cell(rowTong, 1).Value = "Tổng";
                ws.Cell(rowTong, 1).Style.Font.Bold = true;
                ws.Cell(rowTong, 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

                // Công thức SUM cho cột Giá trị tồn (H)
                var cellTongGiaTri = ws.Cell(rowTong, 8);
                cellTongGiaTri.FormulaA1 = $"=SUM(H2:H{rowTong - 1})";
                cellTongGiaTri.Style.Font.Bold = true;
                cellTongGiaTri.Style.NumberFormat.Format = "#,##0";
                cellTongGiaTri.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;

                // Kẻ viền cho toàn bộ bảng tính
                var rangeAll = ws.Range(1, 1, rowTong, headers.Length);
                rangeAll.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                rangeAll.Style.Border.InsideBorder = XLBorderStyleValues.Thin;

                // AutoFilter cho vùng dữ liệu (từ dòng 1 đến dòng dữ liệu cuối, không tính dòng Tổng)
                if (rowTong > 2)
                {
                    ws.Range(1, 1, rowTong - 1, headers.Length).SetAutoFilter();
                }

                // Cố định dòng tiêu đề (Freeze)
                ws.SheetView.FreezeRows(1);

                // Tự động căn chỉnh độ rộng cột
                ws.Columns().AdjustToContents();

                workbook.SaveAs(sfd.FileName);

                if (MessageBox.Show("Xuất file Excel tồn kho thành công! Bạn có muốn mở file không?", "Xác nhận", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                {
                    Process.Start(new ProcessStartInfo(sfd.FileName) { UseShellExecute = true });
                }
            }
            catch (IOException)
            {
                MessageBox.Show("File Excel đang được mở bởi ứng dụng khác. Vui lòng đóng file và thử lại!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi khi xuất Excel tồn kho: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}