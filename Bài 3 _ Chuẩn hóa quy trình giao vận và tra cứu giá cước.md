1\. Phân tích: Mapping VLOOKUP vào "Nhân viên kho tra hồ sơ"  
lookup\_value: Mã vận đơn cầm trên tay (thông tin gốc để đối chiếu).  
table\_array: Tủ hồ sơ chứa bảng giá cước (nơi tra cứu, cột đầu tiên chứa mã vận đơn).  
col\_index\_num: Số thứ tự của cột "Giá cước" tính từ trái sang trong tủ hồ sơ.  
range\_lookup: Số 0 (hoặc FALSE) để yêu cầu nhân viên tìm chính xác tuyệt đối mã vận đơn.  
2\. Triển khai & Ràng buộc cho FastMove  
Excel: Công thức tại ô đầu tiên có dạng \=VLOOKUP(A2, E\$ 2:G\$100, 3, 0). Bắt buộc dùng dấu \$ để cố định vùng bảng tra cứu khi kéo xuống 500 dòng. Tuyệt đối không gộp ô (Merge Cell) ở vùng dữ liệu.  
Word: Thiết kế quy trình bằng bộ Styles (Heading, Body) và đổi màu chữ/bảng theo đúng màu sắc thương hiệu FastMove.  
PowerPoint: Mỗi slide chỉ để 3–4 bullet points (Định luật Miller). Tiêu đề phải là câu kết luận/hành động  
3\. Kiểm thử: Xử lý lỗi \#N/A bằng IFERROR  
Khi mã vận đơn chứa ký tự lạ, lồng VLOOKUP vào trong IFERROR để ẩn lỗi:  
\=IFERROR(VLOOKUP(A2, \$ E\$ 2:\$ G\$ 100, 3, 0), "Kiểm tra lại mã") (hoặc trả về 0).  
