# CY_UnrealHeaderAddin
Excel Addin to help create 'csv' and 'header' files

해당 코드는 "이용섭" & "박청아" 의 공동 소유물이며 허락없이 배포 및 상업적 용도로 사용시
법적 제제 대상이 될 수 있음을 명시하십시오.


코드 사용 및 테이블 정의 규칙

1. 테이블 제작 규칙
   -  첫번 째 열은 변수 이름으로 설정해야 한다. (필수)
   -  두번 째 열은 변수의 데이터 형식으로 설정해야 한다 (필수)
   -  데이터 형식 예시 : int32, float, FString, FName, TArray<int32>, TArray<FString>

2. 파일 사용 규칙
   -  해당 코드는 통째로 특정 경로에 설정해 두어야 한다(권고)
   -  해당 파일 권장 경로  :  C:\Users\"사용자"\source\repos\(해당 파일 묶음)
   -  엑셀을 켠 상태에서 도구 추가모음에서 CY_UnrealHeaderAddin 찾은 후 체크박스 체크

3. Generate Unreal Project (update. 24/03/15)
   - 등록해 둔 테이블 폴더가 해당 프로젝트 내에 위치해 있어야한다.
	(ProjectFolder)\Contents\TableData
   - 엔진 파일의 경로는 정적으로 배치를 시켜두어서
	정해둔 위치가 아니면 제너레이트를 도와주는 툴을 찾을 수 없다
	저장해둔 경로
	1) C:\Program Files (x86)\Epic Games\UE_5.1
	2) C:\Program Files\Epic Games\UE_5.1
	3) D:\UE_5.1
