let danhSachThiSinh = [];

// 1. Tự động tải và đọc file Excel khi mở trang web
window.addEventListener('DOMContentLoaded', async () => {
    try {
        // Thêm ?v=3 để trình duyệt luôn tải bản Excel mới nhất
        let response = await fetch('DANH_SACH_THI_SINH  tra cuu.xlsx?v=3');
        let arrayBuffer = await response.arrayBuffer();
        
        let workbook = XLSX.read(arrayBuffer, { type: 'array' });
        let firstSheetName = 'Danh_Sach_Thi_Sinh'; 
        let worksheet = workbook.Sheets[firstSheetName] || workbook.Sheets[workbook.SheetNames[0]];
        
        // Đọc dữ liệu thô (raw: true) để lấy đúng số ngày Excel và giữ số 0 ở đầu CCCD / Số GPLX
        danhSachThiSinh = XLSX.utils.sheet_to_json(worksheet, { raw: true });
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

            // So sánh khớp cả trường hợp có hoặc mất số 0 ở đầu
            let ketQua = danhSachThiSinh.find(item => {
                let cmt = String(item['SO_CMT'] || '').trim();
                return cmt === soCanCuoc || cmt.padStart(12, '0') === soCanCuoc.padStart(12, '0');
            });

            if (ketQua) {
                hienThiKetQua(ketQua);
            } else {
                hienThiKhongTimThay();
            }
        });
    }
});

// Hàm đổi ngày tháng từ Excel (dùng chung cho cả Ngày sinh và Ngày hết hạn GPLX)
function dinhDangNgayThang(ngay) {
    if (!ngay) return 'Không có';
    
    // Trường hợp 1: Nếu Excel lưu dạng số serial (như 36842, 34345...)
    let num = Number(ngay);
    if (!isNaN(num) && num > 1000) {
        let utcDays = Math.floor(num - 25569);
        let dateInfo = new Date(utcDays * 86400 * 1000);
        let day = String(dateInfo.getUTCDate()).padStart(2, '0');
        let month = String(dateInfo.getUTCMonth() + 1).padStart(2, '0');
        let year = dateInfo.getUTCFullYear();
        return `${day}/${month}/${year}`;
    }

    // Trường hợp 2: Nếu là đối tượng Date
    if (ngay instanceof Date && !isNaN(ngay.getTime())) {
        let day = String(ngay.getDate()).padStart(2, '0');
        let month = String(ngay.getMonth() + 1).padStart(2, '0');
        let year = ngay.getFullYear();
        return `${day}/${month}/${year}`;
    }
    
    // Trường hợp 3: Nếu là chuỗi chữ (VD: "1994-01-11" hoặc "11/01/1994")
    let strNgay = String(ngay).trim();
    if (strNgay.includes('-')) {
        let parts = strNgay.split('T')[0].split('-');
        if (parts.length === 3 && parts[0].length === 4) {
            return `${parts[2]}/${parts[1]}/${parts[0]}`;
        }
    }
    
    return strNgay;
}

// Hàm hiển thị đầy đủ thông tin thí sinh kèm theo lưu ý
function hienThiKetQua(data) {
    const container = document.querySelector('#result-container');
    
    let maDK = data['MA_DK'] || 'N/A';
    let hoTen = data['HO_VA_TEN'] || 'N/A';
    let gioiTinh = data['GIOI_TINH'] || 'N/A';
    let ngaySinh = dinhDangNgayThang(data['NGAY_SINH']); 
    let soCMT = String(data['SO_CMT'] || 'N/A').trim();
    let diaChi = data['DIA_CHI_DAY_DU'] || 'N/A';
    let soBaoDanh = data['SO_BAO_DANH'] || 'N/A';
    let noiDungSH = data['NOI_DUNG_SH'] || 'N/A';
    let tenFileAnh = data['TEN_FILE_ANH'] || '';

    // 3 cột thông tin GPLX đã có vừa bổ sung
    let soGPLXDaCo = data['SO_GPLX_DA_CO'] ? String(data['SO_GPLX_DA_CO']).trim() : '