document.getElementById("callApi").addEventListener("click", function () {
  const sheetNames = [
    "[🥇Performance] Keywords",
    "💎Performance THAY MÀN HÌNH",
    "💎Performance THAY PIN",
    "💎Performance THAY MẶT KÍNH",
    "💎Performance THAY KÍNH LƯNG",
    "💎Performance SỬA ĐIỆN THOẠI",
    "AUDIT URL DV Push SEO",
  ]; // Example, can be dynamic
  const spreadsheetId = "1xXv5NWRtbtnk85wN9_lpxWIqvQLA4IxgcGW7uszB45I"; // Replace with your actual sheet ID

  fetch("http://localhost:3000/update-ranking-and-notes", {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
    },
    body: JSON.stringify({
      sheetNames: sheetNames,
      spreadsheetId: spreadsheetId,
    }),
  })
    .then((response) => response.json())
    .then((data) => {
      console.log("Success:", data);
    })
    .catch((error) => {
      console.error("Error:", error);
    });
});
