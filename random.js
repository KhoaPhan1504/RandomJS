const fs = require('fs');
const XLSX = require('xlsx');

function generateRandomName() {
    const ho = ["Nguyễn", "Trần", "Lê", "Phạm", "Hoàng", "Huỳnh", "Phan", "Vũ", "Võ", "Đặng", "Tạ", "Tòng", "Lâm", "Lý", "Hà"];
    const dem = ["Thị", "Văn", "Công", "Đức", "Hữu", "Tuấn", "Minh", "Hải", "Đức", "Văn", "Hà", "Anh", "Bình", "Ánh", "Tiến", "Thu", "Trọng"];
    const ten = ["An", "Bình", "Cường", "Duy", "Hà", "Hải", "Hòa", "Hùng", "Linh", "Tâm", "Ngân", "Hoàng", "Nam", "Khoa", "Huyền", "Duyên", "Phong", "Thông", "Thắng", "Hường", "Thuý"];
    return `${ho[Math.floor(Math.random() * ho.length)]} ${dem[Math.floor(Math.random() * dem.length)]} ${ten[Math.floor(Math.random() * ten.length)]}`;
}

function generateRandomPhoneNumber() {
    let phoneNumber = '0';
    for (let i = 0; i < 9; i++) {
        phoneNumber += Math.floor(Math.random() * 10);
    }
    return phoneNumber;
}

function generateRandomAddress() {
    const ten = ["An", "Bình", "Cường", "Duy", "Hà", "Hải", "Hòa", "Hùng", "Linh", "Tâm", "Tiến", "Long", "Lâm", "Huyền"];
    return `Số ${Math.floor(Math.random() * 200) + 1}, Đường ${ten[Math.floor(Math.random() * ten.length)]}, Quận/Huyện ${ten[Math.floor(Math.random() * ten.length)]}, Tỉnh/Thành phố ${ten[Math.floor(Math.random() * ten.length)]}`;
}

function generateRandomEmail() {
    const ten = ["An", "Bình", "Cường", "Duy", "Hà", "Hải", "Hòa", "Hùng", "Linh", "Tâm", "Tiến", "Long", "Lâm", "Huyền"];
    const domains = ["gmail.com", "yahoo.com", "hotmail.com", "outlook.com", "edu.vn"];
    return `${ten[Math.floor(Math.random() * ten.length)]}${Math.floor(Math.random() * 100)}@${domains[Math.floor(Math.random() * domains.length)]}`;
}

// Tạo mảng chứa dữ liệu
const data = [["Họ và Tên", "Số điện thoại", "Địa chỉ", "Email"]];
for (let i = 0; i < 200; i++) {
    data.push([generateRandomName(), generateRandomPhoneNumber(), generateRandomAddress(), generateRandomEmail()]);
}

// Tạo workbook và ghi dữ liệu vào sheet
const workbook = XLSX.utils.book_new();
const worksheet = XLSX.utils.aoa_to_sheet(data);
XLSX.utils.book_append_sheet(workbook, worksheet, "Danh sách người");

// Lưu file Excel
const excelFileName = "DanhSachNguoi.xlsx";
XLSX.writeFile(workbook, excelFileName);
console.log(`File Excel đã được tạo thành công: ${excelFileName}`);
