# CY_UnrealHeaderAddin
Excel Addin to help create 'csv' and 'header' files

해당 코드는 "이용섭" & "박청아" 의 공동 소유물이며 허락없이 배포 및 상업적 용도로 사용시
법적 제제 대상이 될 수 있음을 명시하십시오.


코드 사용 및 테이블 정의 규칙

1. 테이블 제작 규칙
   -  첫번 째 열은 변수 이름으로 설정해야 한다. (필수)
   -  두번 째 열은 변수의 데이터 형식으로 설정해야 한다 (필수)
   -  데이터 형식 예시 : int32, float, FString, FName, TArray<int32>, TArray<FString>

2. 코드 사용 규칙
   -  해당 코드는 통째로 특정 경로에 설정해 두어야 한다(권고)
   -  권장 경로  :  C:\Users\"사용자"\source\repos\(해당 파일 묶음)
   -  엑셀을 켠 상태에서 도구 추가모음에서 CY_UnrealHeaderAddin 찾은 후 체크박스 체크
