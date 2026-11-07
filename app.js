let danhSachThiSinh = [];
let daTaiXong = false;

// 1. Tự động tìm và tải file Excel khi mở trang web
window.addEventListener('DOMContentLoaded', async () => {
    const danhSachTenFile = [
        'DANH_SACH_THI_SINH  tra cuu.xlsx',
        encodeURI('DANH_SACH_THI_SINH  tra cuu.xlsx'),
        'DANH_SACH_THI_SINH tra cuu.xlsx',
        'mon 2.xlsx'
    ];

    let response = null;
    for (let tenFile of danhSachTenFile) {
        try {
            let res = await fetch(tenFile + '?t=' + Date.now());
            if (res.ok) {
                response = res;
                console.log("Đã kết nối file:", tenFile);
                break;
            }
        } catch (err) {}
    }

    if (!response) {
        const container = document.querySelector('#result-container');
        if (container) {
            container.innerHTML = `
                <div style="background: #fff5f5; padding: 16px; border-radius: 8px; border: 1px solid #feb2b2; color: #c53030; text-align: center;">
                    <strong>Chưa tìm thấy file Excel dữ liệu!</strong><br>Vui lòng kiểm tra xem file <code>DANH_SACH_THI_SINH  tra cuu.xlsx</code> đã được tải lên GitHub chưa.
                </div>
            `;
        }
        return;
    }

    try {
        let arrayBuffer = await response.arrayBuffer();
        let workbook = XLSX.read(arrayBuffer, { type: 'array' });
        let firstSheetName = workbook.SheetNames.includes('Danh_Sach_Thi_Sinh') 
            ? 'Danh_Sach_Thi_Sinh' 
            : workbook.SheetNames[0];
        let worksheet = workbook.Sheets[firstSheetName];
        
        danhSachThiSinh = XLSX.utils.sheet_to_json(worksheet, { raw: true });
        daTaiXong = true;
        console.log("Đã tải xong:", danhSachThiSinh.length, "hồ sơ.");
    } catch (error) {
        console.error("Lỗi đọc Excel:", error);
    }
});

// 2. Xử lý khi bấm nút Tra cứu
document.addEventListener('DOMContentLoaded', () => {
    const lookupForm = document.querySelector('#lookup-form');
    const inputCCCD = document.querySelector('#identity-number');

    if (lookupForm && inputCCCD) {
        lookupForm.addEventListener('submit', (e) => {
            e.preventDefault();
            
            if (!daTaiXong) {
                alert("Hệ thống đang tải dữ liệu, vui lòng đợi 2 giây rồi thử lại!");
                return;
            }

            let soCanCuoc = inputCCCD.value.trim();
            if (!soCanCuoc) {
                alert("Vui lòng nhập số căn cước công dân!");
                return;
            }

            let ketQua = danhSachThiSinh.find(item => {
                let cmt = String(item['SO_CMT'] || item['Số căn cước'] || '').trim();
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

// Hàm đổi ngày tháng từ Excel sang DD/MM/YYYY chuẩn xác
function dinhDangNgayThang(ngay) {
    if (!ngay) return 'Không có';
    
    let num = Number(ngay);
    if (!isNaN(num) && num > 1000) {
        let utcDays = Math.floor(num - 25569);
        let dateInfo = new Date(utcDays * 86400 * 1000);
        let day = String(dateInfo.getUTCDate()).padStart(2, '0');
        let month = String(dateInfo.getUTCMonth() + 1).padStart(2, '0');
        let year = dateInfo.getUTCFullYear();
        return `${day}/${month}/${year}`;
    }

    if (ngay instanceof Date && !isNaN(ngay.getTime())) {
        let day = String(ngay.getDate()).padStart(2, '0');
        let month = String(ngay.getMonth() + 1).padStart(2, '0');
        let year = ngay.getFullYear();
        return `${day}/${month}/${year}`;
    }
    
    let strNgay = String(ngay).trim();
    if (strNgay.includes('-')) {
        let parts = strNgay.split('T')[0].split('-');
        if (parts.length === 3 && parts[0].length === 4) {
            return `${parts[2]}/${parts[1]}/${parts[0]}`;
        }
    }
    
    return strNgay;
}

// Hàm hiển thị thông tin thí sinh
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

    let soGPLXDaCo = data['SO_GPLX_DA_CO'] ? String(data['SO_GPLX_DA_CO']).trim() : 'Không có';
    let hangGPLXDaCo = data['HANG_GPLX_DA_CO'] ? String(data['HANG_GPLX_DA_CO']).trim() : 'Không có';
    let ngayHHGPLXDaCo = data['NGAY_HH_GPLX_DACO'] ? dinhDangNgayThang(data['NGAY_HH_GPLX_DACO']) : 'Không có';

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
                        <p><strong>Ngày sinh:</strong> ${ngaySinh}</p>
                        <p><strong>Giới tính:</strong> ${gioiTinh}</p>
                        <p><strong>Số CCCD/CMT:</strong> ${soCMT}</p>
                        <p><strong>Số báo danh:</strong> <span style="color: #c53030; font-weight: bold;">${soBaoDanh}</span></p>
                        <p><strong>Nội dung sát hạch:</strong> ${noiDungSH}</p>
                        <p><strong>Số GPLX đã có:</strong> ${soGPLXDaCo}</p>
                        <p><strong>Hạng GPLX đã có:</strong> ${hangGPLXDaCo}</p>
                        <p><strong>Ngày hết hạn GPLX đã có:</strong> ${ngayHHGPLXDaCo}</p>
                        <p><strong>Địa chỉ:</strong> ${diaChi}</p>
                    </div>
                </div>

                <div style="margin-top: 16px; padding: 12px 16px; background-color: #fff5f5; border-left: 4px solid #e53e3e; border-radius: 4px; font-size: 0.95rem; color: #742a2a; line-height: 1.6;">
                    <strong>Lưu ý:</strong> nếu thông tin bị sai, vui lòng nhắn Zalo lại cho Nhóm Trưởng trong giờ hành chính để chỉnh sửa lại trước ngày thi. Nếu không chỉnh sửa, đến ngày thi sẽ không sửa được. Cảm ơn.
                </div>
            </div>
        `;
    }
}

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