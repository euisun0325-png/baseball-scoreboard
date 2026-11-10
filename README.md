# baseball-scoreboard

야구 스코어보드 안드로이드 앱 업데이트 배포용 저장소예요.

- `version.json` : 최신 버전 정보 (앱이 켜질 때 이 파일을 확인)
- `BaseballScoreboard_v버전.apk` : 최신 앱 설치 파일 (버전마다 이름이 달라 예전 파일이 섞이지 않아요)

v6.3부터 앱이 이 저장소를 확인해 새 버전이 있으면 알려 주고 바로 설치할 수 있어요.

## 욜로 화면전환 앱

- `remote/version.json` : 최신 버전 정보
- `remote/YoloRemote_v버전.apk` : 최신 앱 설치 파일

v1.4부터 같은 방식으로 업데이트 알림이 떠요.

## 가입·로그인 서버

- `server/Code.gs` : 구글 시트에 붙여 넣는 가입·승인 서버 코드
- `config.json` 의 `auth` 에 서버 주소를 넣으면 앱에서 로그인이 켜져요 (비우면 꺼짐)
