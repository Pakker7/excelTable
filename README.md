# Excel Table
> excelTable.js는 엑셀 데이터 결과를 엑셀표 모양으로 출력하는 자바스크립트 라이브러리 
> 원본 엑셀을 출력하고, 에러난 부분을 추가로 표시한다.

## 사용 방법
- 웹 페이지에서 사용하려면 hangul.js 파일을 <script>태그를 이용하여 추가
- bootstrap 및 jquery 가 필요함
	```	<link rel="stylesheet" href="./excelTable.css"/>
		<script src="excelTable.js"></script>
	``` 
- 작성 방법
  - target or targetObj
    - 필수 값
    - **target은 id명**을 입력
    - **targetObj는 객체**를 입력
  - visibleUnique
    -  유니크 키 표시 유무
  - data
    - 배열로 전달 시
      - 배열 형태로 전달 한다.
    - 객체로 전달 시 
	    - 필수 값
	    - header
	      - 엑셀 상위의 컬럼명
	      - displayName 출력되는 컬럼 명
	      - columnName 값이 origin의 key값과 연결되어 출력
	      - dataType NUMBER or STRING 에 따라 왼쪽, 오른쪽 정렬 됨
	      - unique 상위의 visibleUnique가 true일 때 컬럼 왼편에 키 표기
	    - origin
	      - 원본 데이터
	    - errors
	      - 에러난 컬럼 명시
	      - 에러난 컬럼이 존재하는 행과 열의 폰트 색을 바꿔줌
  - style
    - 미작성시 default 스타일로 지정(생략가능)
    - fontColor, fontSize, backgroundColor, warnColor(error) 커스텀 가능
```
	excelTable.init({
		//target : "targetTable", 						
		targetObj : $("#targetTable").get(0),
		visibleUnique : false, 				
		//data : [
		//	["네이버","7","그룹1","그룹2","그룹3"], 
		//	["다음","10","그룹1","그룹2","그룹3","도메인"], 
		//	["구글","15","그룹1","그룹2","그룹3","도메인"], 
		//],
		data : {
			header : header,
			origin : origin,
			errors : errors,
		},
		style : {
			edge : {
				fontColor : 'green', fontSize : '20px'
			},
			header : {
				backgroundColor  : 'green', fontColor : 'white'
			},
			cell : {
				backgroundColor  : 'blue' , fontColor : 'yellow'
			},
			error : {
				warnColor : 'grey', fontColor : 'white'
			}
		}		
	});
```

## 예시
- index.html 참고
