# Send-Email
### win32com-pandas-Project
- 2022-11-30 시작!🏃‍♀️
---
## 🎉 프로젝트 소개
#### 😣 똑같은 이메일 작업은 지겨워! 😭
- 자동 메일발송 시스템으로 반복적인 이메일 발송⚡을 빠르게 처리해보자
- 클릭 몇번으로 **1시간짜리 일을 👉 단 1분만에 해결!**
<br>


## ❓ 프로젝트 계기
- 어느날 영업관리팀에서 날아온 **HELP!**요청
- _**막내 사원이 업체에 메일보내는데 시간을 너무 뺏겨요😪 메일발송 자동화 시스템을 개발해주세요!**_


#### <작업 요청사항>
1. Xlsx 파일 읽어 각 column의 내용을 parsing
2. Parsing된 내용에서 첫번째 칼럼인 센터명을 통해 메일 제목에 센터이름 추가
3. 메일 제목에 현재날짜 포함
4. 메일 제목에 "출고건", "결과발송건" 구분하여 표기
5. 메일 본문에 고정된 멘트 추가
6. 고정된 멘트 사이에 xlsx 파일 내용을 표시 (아래 그림 참조)
7. Xlsx 파일에는 1개의 센터만 있는 것이 아니라, 여러 센터가 있으므로 센터별로 구별하여 메일을 발송해야 함.
<br>
<img width="40%" src="https://user-images.githubusercontent.com/118783464/207486993-f689830a-f8dc-4e75-9017-d27bc9ec763a.png"/>
<br>

- ~~그리하여 개발 시작...🤣🤣~~
<br>


## ✨ 주요 기능
- **Pandas**로 Xlsx 파일 읽어 데이터프레임화 📋
- **HTML** 형식으로 메일 본문 내용 작성 📝
- **Wincom32.client** 라이브러리 사용하여 메일 발송!📤

