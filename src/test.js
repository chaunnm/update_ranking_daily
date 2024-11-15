document.getElementById("callApi").addEventListener("click", function () {
  const sheetNames = [
    "[🥇Performance] Keywords",
    "💎Performance THAY MÀN HÌNH",
    "💎Performance THAY PIN",
    "💎Performance THAY MẶT KÍNH",
    "💎Performance THAY KÍNH LƯNG",
    "💎Performance SỬA ĐIỆN THOẠI",
    "AUDIT URL DV Push SEO",
  ]; // Replace with actual sheet names
  const spreadsheetId = "1xXv5NWRtbtnk85wN9_lpxWIqvQLA4IxgcGW7uszB45I"; // Replace with your actual sheet ID

  // Khởi tạo danh sách các promises cho fetch
  const fetchPromises = sheetNames.map((sheet) => {
    const startTime = Date.now(); // Ghi lại thời gian bắt đầu
    return fetch("http://localhost:3000/update-ranking-and-notes", {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
      },
      body: JSON.stringify({
        sheetName: sheet,
        spreadsheetId: spreadsheetId,
      }),
    })
      .then(async (res) => {
        const data = await res.json();
        const endTime = Date.now(); // Ghi lại thời gian kết thúc
        const duration = (endTime - startTime) / 1000; // Tính thời gian thực thi (giây)
        console.log(`Sheet: ${sheet} completed in ${duration} seconds.`);
        return { sheet, success: true, duration, data }; // Trả về kết quả thành công
      })
      .catch((error) => {
        const endTime = Date.now(); // Ghi lại thời gian kết thúc
        const duration = (endTime - startTime) / 1000; // Tính thời gian thực thi
        console.error(`Error for sheet: ${sheet}`, error);
        return { sheet, success: false, duration, error }; // Trả về kết quả lỗi
      });
  });

  // Chạy tất cả fetch song song với Promise.all
  Promise.all(fetchPromises).then((results) => {
    console.log("All requests completed.");
    results.forEach((result) => {
      if (result.success) {
        console.log(
          `Sheet: ${result.sheet}, Duration: ${result.duration} seconds, Response:`,
          result.data
        );
      } else {
        console.log(
          `Sheet: ${result.sheet}, Duration: ${result.duration} seconds, Error:`,
          result.error
        );
      }
    });
  });
});
