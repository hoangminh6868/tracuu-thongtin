let danhSachThiSinh = [];

// 1. Tự động tải và đọc file Excel khi mở trang web
window.addEventListener('DOMContentLoaded', async () => {
    try {
        let response = await fetch('DANH_SACH_THI_SINH  tra cuu.xlsx');
        let arrayBuffer = await response.arrayBuffer();
        
        let workbook = XLSX.read(arrayBuffer, { type: 'array', cellDates: true }); // cellDates: true giúp đọc đúng ngày tháng
        let firstSheetName = 'Danh_Sach_Thi_Sinh'; 
        let worksheet = workbook.Sheets[firstSheetName] || workbook.Sheets[workbook.SheetNames[0]];
        
        danhSachThiSinh = XLSX.utils.sheet_to_json(worksheet);
        console.log("Đã tải dữ liệu thành công:", danhSachThiSinh.length, "bản ghi.");
    } catch (error) {
        console.error("Lỗi khi đọc file Excel:", error);
    }
});

// 2. Xử lý sự kiện khi người dùng bấm Tra cứu
document.addEventListener('DOMContentLoaded', () => {
    const lookupForm = document.querySelector('#lookup-form');
    const inputCCCD = document.querySelector('#identity-number');

    if (lookupForm && inputCCCD) {
        lookupForm.addEventListener('submit', (e) => {
            e.preventDefault(); 
            let soCanCuoc = inputCCCD.value.trim();

            if (!soCanCuoc) {
                alert("Vui lòng nhập số căn cước công dân!");
                return;
            }

            let ketQua = danhSachThiSinh.find(item => 
                String(item['SO_CMT'] || '').trim() === soCanCuoc
            );

            if (ketQua) {
                hienThiKetQua(ketQua);
            } else {
                hienThiKhongTimThay();
            }
        });
    }
});

// Hàm định dạng ngày sinh chuẩn xác, hỗ trợ mọi kiểu dữ liệu từ Excel
function dinhDangNgaySinh(ngay) {
    if (!ngay) return 'N/A';
    
    // Nếu là đối tượng Date chuẩn từ Excel
    if (ngay instanceof Date && !isNaN(ngay.getTime())) {
        let day = String(ngay.getDate()).padStart(2, '0');
        let month = String(ngay.getMonth() + 1).padStart(2, '0');
        let year = ngay.getFullYear();
        return `${day}/${month}/${year}`;
    }
    
    // Nếu là chuỗi (ví dụ dạng "2005-11-11" hoặc "11/11/2005")
    let strNgay = String(ngay).trim();
    if (strNgay.includes('-')) {
        let parts = strNgay.split('T')[0].split('-'); // Tách lấy phần YYYY-MM-DD
        if (parts.length === 3) {
            return `${parts[2]}/${parts[1]}/${parts[0]}`;
        }
    }
    
    return strNgay;
}

// Hàm hiển thị thông tin thí sinh kèm theo lưu ý
function hienThiKetQua(data) {
    const container = document.querySelector('#result-container');
    
    let hoTen = data['HO_VA_TEN'] || 'N/A';
    let ngaySinhRaw = data['NGAY_SINH'] || 'N/A';
    let ngaySinh = dinhDangNgaySinh(ngaySinhRaw); // Hiển thị đúng định dạng DD/MM/YYYY
    let gioiTinh = data['GIOI_TINH'] || 'N/A';
    let soCMT = data['SO_CMT'] || 'N/A';
    let diaChi = data['DIA_CHI_DAY_DU'] || 'N/A';
    let soBaoDanh = data['SO_BAO_DANH'] || 'N/A';
    let noiDungSH = data['NOI_DUNG_SH'] || 'N/A';
    let maDK = data['MA_DK'] || 'N/A';
    let tenFileAnh = data['TEN_FILE_ANH'] || '';

    if (container) {
        container.innerHTML = `
            <div style="background: #fff; padding: 24px; border-radius: 12px; box-shadow: 0 4px 12px rgba(0,0,0,0.08); border: 1px solid #fed7d7;" class="space-y-4">
                <h3 style="color: #9d1d2b; font-size: 1.25rem; font-weight: bold; border-bottom: 2px solid #fed7d7; padding-bottom: 8px; margin-bottom: 16px;">Thông tin hồ sơ thí sinh</h3>
                
                <div style="display: flex; flex-wrap: wrap; gap: 20px; align-items: flex-start;">
                    ${tenFileAnh ? `
                    <div style="text-align: center; flex: 0 0 130px;">
                        <img src="images/${tenFileAnh}" alt="Ảnh thí sinh" style="width: 120px; height: 150px; object-fit: cover; border-radius: 6px; border: 1px solid #d1d5db;" onerror="this.style.display='none'" />
                    </div>` : ''}
                    
                    <div style="flex: 1; min-width: 240px; line-height: 1.8;">
                        <p><strong>Mã đăng ký:</strong> ${maDK}</p>
                        <p><strong>Họ và tên:</strong> <span style="color: #1d4ed8; font-weight: 700; font-size: 1.1rem;">${hoTen}</span></p>
                        <p><strong>Ngày sinh:</strong> ${ngaySinh} &nbsp;|&nbsp; <strong>Giới tính:</strong> ${gioiTinh}</p>
                        <p><strong>Số CCCD/CMT:</strong> ${soCMT}</p>
                        <p><strong>Số báo danh:</strong> <span style="color: #c53030; font-weight: bold;">${soBaoDanh}</span></p>
                        <p><strong>Nội dung sát hạch:</strong> ${noiDungSH}</p>
                        <p><strong>Địa chỉ:</strong> ${diaChi}</p>
                    </div>
                </div>

                <!-- Phần lưu ý bổ sung -->
                <div style="margin-top: 16px; padding: 12px 16px; background-color: #fff5f5; border-left: 4px solid #e53e3e; border-radius: 4px; font-size: 0.95rem; color: #742a2a; line-height: 1.6;">
                    <strong>Lưu ý:</strong> nếu thông tin bị sai, vui lòng nhắn Zalo lại cho Nhóm Trưởng trong giờ hành chính để chỉnh sửa lại trước ngày thi. Nếu không chỉnh sửa, đến ngày thi sẽ không sửa được. Cảm ơn.
                </div>
            </div>
        `;
    }
}

// Hàm hiển thị thông báo khi không tìm thấy số CCCD
function hienThiKhongTimThay() {
    const container = document.querySelector('#result-container');
    if (container) {
        container.innerHTML = `
            <div style="background: #fff5f5; padding: 20px; border-radius: 12px; border: 1px solid #feb2b2; text-align: center; color: #c53030;">
                <h3 style="font-weight: bold; margin-bottom: 8px;">Không tìm thấy hồ sơ</h3>
                <p>Không có thông tin thí sinh nào ứng với số căn cước công dân này. Vui lòng kiểm tra lại!</p>
            </div>
        `;
    }
}