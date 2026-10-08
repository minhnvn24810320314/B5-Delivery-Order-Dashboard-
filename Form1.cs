using System;
using System.Drawing;
using System.Windows.Forms;

namespace NGUYEN_VU_NHAT_MINH_24810320314_BAI_TAP_5
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            // Thiết lập giá trị ban đầu cho ComboBox
            if (cboLoaiVC.Items.Count > 0)
            {
                cboLoaiVC.SelectedIndex = 0;
            }

            // Sinh mã đơn hàng tự động
            txtMaDon.Text = "DH-" + DateTime.Now.ToString("yyyyMMdd-HHmm");
            txtTenKH.Text = "Nguyễn Vũ Nhật Minh";
            txtSoDienThoai.Text = "0987654321";
            txtDiaChi.Text = "99 Nguyễn Huệ, Q.1, TP.HCM";

            // Nạp một số dòng hàng hóa mẫu
            NapDuLieuMau();

            // Cập nhật ngay thời gian và thanh StatusStrip
            CapNhatThoiGianSystem();
            CapNhatStatusStripRealtime();
        }

        private void NapDuLieuMau()
        {
            dgvHangHoa.Rows.Add("Màn hình Dell UltraSharp 27\"", 2, 4.5, 8500000, 17000000);
            dgvHangHoa.Rows.Add("Bàn phím cơ Logitech G Pro", 3, 1.2, 2500000, 7500000);
            dgvHangHoa.Rows.Add("Chuột không dây MX Master 3S", 5, 0.8, 2200000, 11000000);
        }

        #region XỬ LÝ PHÍM TẮT DÙNG CHUNG (F2 & DELETE)
        private void Form1_KeyDown(object sender, KeyEventArgs e)
        {
            // Phím F2: Thêm dòng mới vào DataGridView
            if (e.KeyCode == Keys.F2)
            {
                btnThemDong_Click(sender, e);
                e.Handled = true;
            }
            // Phím Delete: Xóa dòng đang chọn (khi không trong chế độ edit cell)
            else if (e.KeyCode == Keys.Delete && !dgvHangHoa.IsCurrentCellInEditMode)
            {
                btnXoaDong_Click(sender, e);
                e.Handled = true;
            }
        }
        #endregion

        #region THAO TÁC THÊM / XÓA DÒNG DATAGRIDVIEW
        private void btnThemDong_Click(object sender, EventArgs e)
        {
            // Thêm 1 dòng trống mặc định
            int rowIndex = dgvHangHoa.Rows.Add("", 1, 1.0, 0, 0);
            dgvHangHoa.CurrentCell = dgvHangHoa.Rows[rowIndex].Cells[0];
            dgvHangHoa.BeginEdit(true);

            CapNhatStatusStripRealtime();
        }

        private void btnXoaDong_Click(object sender, EventArgs e)
        {
            if (dgvHangHoa.SelectedRows.Count > 0)
            {
                foreach (DataGridViewRow row in dgvHangHoa.SelectedRows)
                {
                    if (!row.IsNewRow)
                    {
                        dgvHangHoa.Rows.Remove(row);
                    }
                }
            }
            else if (dgvHangHoa.CurrentRow != null && !dgvHangHoa.CurrentRow.IsNewRow)
            {
                dgvHangHoa.Rows.Remove(dgvHangHoa.CurrentRow);
            }
            else
            {
                MessageBox.Show("Vui lòng chọn dòng hàng hóa cần xóa!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            CapNhatStatusStripRealtime();
        }
        #endregion

        #region VALIDATION CELL & ERRORPROVIDER (CẢNH BÁO SỐ LƯỢNG / TRỌNG LƯỢNG <= 0)
        private void dgvHangHoa_CellValidating(object sender, DataGridViewCellValidatingEventArgs e)
        {
            // Kiểm tra cột Số lượng (cột 1) và Trọng lượng (cột 2)
            if (e.ColumnIndex == 1 || e.ColumnIndex == 2)
            {
                string columnName = dgvHangHoa.Columns[e.ColumnIndex].HeaderText;
                string valStr = e.FormattedValue.ToString();

                if (double.TryParse(valStr, out double val))
                {
                    if (val <= 0)
                    {
                        dgvHangHoa.Rows[e.RowIndex].ErrorText = $"{columnName} phải lớn hơn 0!";
                        MessageBox.Show($"{columnName} tại dòng {e.RowIndex + 1} phải lớn hơn 0!",
                                        "Cảnh báo dữ liệu (ErrorProvider)", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        e.Cancel = true; // Chặn rời khỏi cell khi nhập sai
                    }
                    else
                    {
                        dgvHangHoa.Rows[e.RowIndex].ErrorText = string.Empty;
                    }
                }
                else
                {
                    dgvHangHoa.Rows[e.RowIndex].ErrorText = $"{columnName} phải là số hợp lệ!";
                    MessageBox.Show($"{columnName} tại dòng {e.RowIndex + 1} phải là số!",
                                    "Lỗi định dạng", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    e.Cancel = true;
                }
            }
        }

        private void dgvHangHoa_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            // Tính lại Thành tiền = Số lượng * Đơn giá khi sửa Số lượng (cột 1) hoặc Đơn giá (cột 3)
            if (e.ColumnIndex == 1 || e.ColumnIndex == 3)
            {
                DataGridViewRow row = dgvHangHoa.Rows[e.RowIndex];

                int.TryParse(row.Cells[1].Value?.ToString(), out int soLuong);
                double.TryParse(row.Cells[3].Value?.ToString(), out double donGia);

                double thanhTien = soLuong * donGia;
                row.Cells[4].Value = thanhTien;
            }

            // Cập nhật ngay các con số trên thanh StatusStrip
            CapNhatStatusStripRealtime();
        }
        #endregion

        #region TÍNH TOÁN VÀ CẬP NHẬT STATUSSTRIP REALTIME
        private void CapNhatStatusStripRealtime()
        {
            int tongSoLuong = 0;
            double tongTrongLuong = 0;
            double tongTien = 0;

            foreach (DataGridViewRow row in dgvHangHoa.Rows)
            {
                if (row.IsNewRow) continue;

                int.TryParse(row.Cells[1].Value?.ToString(), out int sl);
                double.TryParse(row.Cells[2].Value?.ToString(), out double tl);
                double.TryParse(row.Cells[4].Value?.ToString(), out double tt);

                tongSoLuong += sl;
                tongTrongLuong += tl;
                tongTien += tt;
            }

            lblStatusTongSL.Text = $"Tổng số lượng: {tongSoLuong}";
            lblStatusTongTL.Text = $"Tổng trọng lượng: {tongTrongLuong:N1} kg";
            lblStatusTongTien.Text = $"Tổng tiền đơn hàng: {tongTien:N0} VNĐ";
        }

        private void timerSystem_Tick(object sender, EventArgs e)
        {
            CapNhatThoiGianSystem();
        }

        private void CapNhatThoiGianSystem()
        {
            lblStatusTime.Text = "Giờ hệ thống: " + DateTime.Now.ToString("HH:mm:ss dd/MM/yyyy");
        }
        #endregion

        #region LƯU ĐƠN HÀNG VÀ LÀM MỚI FORM
        private void btnLuuDon_Click(object sender, EventArgs e)
        {
            // Validation thông tin khách hàng bằng ErrorProvider
            errorProviderInput.Clear();

            if (string.IsNullOrWhiteSpace(txtMaDon.Text))
            {
                errorProviderInput.SetError(txtMaDon, "Chưa nhập Mã đơn hàng!");
                txtMaDon.Focus();
                return;
            }

            if (string.IsNullOrWhiteSpace(txtTenKH.Text))
            {
                errorProviderInput.SetError(txtTenKH, "Chưa nhập Tên khách hàng!");
                txtTenKH.Focus();
                return;
            }

            if (dgvHangHoa.Rows.Count == 0)
            {
                MessageBox.Show("Danh mục hàng hóa trong đơn trống. Vui lòng thêm ít nhất 1 mặt hàng!",
                                "Thông báo lỗi", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            string thietLapVC = cboLoaiVC.SelectedItem != null ? cboLoaiVC.SelectedItem.ToString() : "";

            string summary = $"LƯU BẢNG ĐIỀU KHIỂN ĐƠN HÀNG THÀNH CÔNG!\n" +
                             $"----------------------------------------\n" +
                             $"• Mã đơn hàng: {txtMaDon.Text.Trim()}\n" +
                             $"• Khách hàng: {txtTenKH.Text.Trim()} ({txtSoDienThoai.Text.Trim()})\n" +
                             $"• Loại vận chuyển: {thietLapVC}\n" +
                             $"• Số mặt hàng: {dgvHangHoa.Rows.Count}\n" +
                             $"• {lblStatusTongSL.Text}\n" +
                             $"• {lblStatusTongTL.Text}\n" +
                             $"• {lblStatusTongTien.Text}";

            MessageBox.Show(summary, "Xác nhận lưu đơn thành công", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void btnLamMoi_Click(object sender, EventArgs e)
        {
            txtMaDon.Text = "DH-" + DateTime.Now.ToString("yyyyMMdd-HHmm");
            txtTenKH.Clear();
            txtSoDienThoai.Clear();
            txtDiaChi.Clear();
            txtGhiChu.Clear();

            if (cboLoaiVC.Items.Count > 0)
                cboLoaiVC.SelectedIndex = 0;

            dgvHangHoa.Rows.Clear();
            errorProviderInput.Clear();

            CapNhatStatusStripRealtime();
            txtTenKH.Focus();
        }
        #endregion
    }
}