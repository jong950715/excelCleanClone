#### It can be used when your excel file is bloating so the file demands exorbitant storage space and works improperly or slow.   
#### Program makes new excel file and copy mandatory data like text.   
##### See below for more detail about what is cloned
1. text (including formula)
2. cell width&height
3. font
    1. font size
    2. font color
    3. cell bg color
    4. bold, italic
    5. ~~font type is excluded~~
   
##### To use this program please check out release page.
##### https://github.com/jong950715/excelCleanClone/releases

---
#### 엑셀이 부풀어서 용량이 과도하고 속도가 느릴 때, 사용하는 툴 입니다.
#### 새로운 엑셀파일을 만들어서 필요한 내용만 복사합니다.
##### 복사내역은 아래와 같습니다.
1. 텍스트 (수식 포함)
2. 셀의 너비와 높이
3. 폰트
    1. 폰트 사이즈
    2. 폰트 색상 (단일 색의 경우만 복사)
    3. 셀 색상
    4. 볼드, 이태릭   
    5. ~~폰트 종류 제외됨.~~
 
#### 릴리즈 페이지를 참고해주세요.
##### https://github.com/jong950715/excelCleanClone/releases
 

##### 사용방법은 아래와 같습니다. (파이썬 설치 되어있다는 가정)

##### 1. 가상환경 만들기
```
python -m venv venv
```

##### 2. 가상환경 실행하기
```
source venv/bin/activate
```
혹은
```
source venv/Scripts/activate
```

##### 3. 라이브러리 다운받기
 ```
pip install req.txt
```

##### 4. 실행하기
```
python main.py -in {파일이름} -X {복사영역 가로} -Y {복사영역 세로}
```
예시
```
python main.py -in input.xlsx -X 50 -Y 1000
```

---
혹은 GUI 개발예정?