// 사랑관 태양광 점검 앱 설정 (여기만 고치면 됩니다)
//  - GAS_URL : 구글 Apps Script 웹 앱 배포 주소(…/exec). 비어 있으면 '연습 모드'(이 기기에만 저장).
//  - MODEL   : 인버터 사진을 읽을 때 쓰는 Claude 모델. (API 키를 넣은 기기에서만 사용)
window.APP_CONFIG = {
  GAS_URL: "https://script.google.com/macros/s/AKfycbyDTbO5yeIziLDFByg8eqGM7_u5lirmlqNdUwfncKt2bnZczAi7Mai0-yo7LEx0HoQihQ/exec",
  MODEL: "claude-sonnet-5-5"
};
