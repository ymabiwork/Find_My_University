Attribute VB_Name = "Module3"
'2025 수능으로 바꾸기와 목표대학 내 점수 확인의 두 가지 일을 동시에 수행하는 것


'2025학년도 수능 점수로 바꾸기 버튼을 눌렀을 때
 
 
 '내 점수 입력 sheet
    '--- 시트 객체 선언 ---
    Dim wsStudentData As Worksheet
    
    '학생의 시험 정보, 변수 이름 sat는 SAT에소 가져옴.
    '모의고사 시험 정보도 저장하는 변수임.
    Dim satYearVal As String  '년도
    Dim satGradeVal As String '학년
    Dim satMonthVal As String '시행 월

    

    '--- 목표 대학 및 지원 정보 변수 선언 ---
    Dim targetUniversity As String
    Dim targetMajor As String
    Dim applicationType As String

    '--- 과목명 정보 변수 선언 ---
'    Dim koreanSubjectName As String
    Dim mathSubjectName As String
    Dim inquiry1SubjectName As String
    Dim inquiry2SubjectName As String
'    Dim secondLanguageSubjectName As String

    '--- 점수 및 수능 구분 정보 변수 선언 ---
    Dim examYearClassification As String
    Dim koreanScore As Double
    Dim mathScore As Double
    Dim englishScore As Double
    Dim koreanHistoryScore As Double
    Dim inquiry1Score As Double
    Dim inquiry2Score As Double
    Dim secondLanguageScore As Double
    
    Dim koreanScoreDisplay As Double
    Dim mathScoreDisplay As Double
    Dim englishScoreDisplay As Double
    Dim koreanHistoryScoreDisplay As Double
    Dim inquiry1ScoreDisplay As Double
    Dim inquiry2ScoreDisplay As Double
    Dim secondLanguageScoreDisplay As Double
    
    
    '----대학 반영 점수 변환 기준 -----
    Dim koreanScoreProcessType  As String
    Dim mathScoreProcessType As String
    Dim englishScoreProcessType As String
    Dim koreanHistoryScoreProcessType As String
    Dim inquiry1ScoreProcessType As String
    Dim inquiry2ScoreProcessType As String
    Dim secondLanguageScoreProcessType As String
    
    
    '변수 할당을 위한 변수가 입력된 주소
    Dim satYearValCellAddress As String  '년도
    Dim satGradeValCellAddress As String '학년
    Dim satMonthValCellAddress As String '시행 월

    Dim targetUniversityCellAddress As String
    Dim targetMajorCellAddress As String
    Dim applicationTypeCellAddress As String
    
    Dim koreanCellAddress As String
    Dim koreanSubjectNameCellAddress As String
    
    Dim mathCellAddress As String
    Dim mathSubjectNameCellAddress As String
    
'    Dim englishNameCellAddress As String
'    Dim koreanHistoryNameCellAddress As String
'    Dim inquiry1SubjectNameCellAddress As String
'    Dim inquiry2SubjectNameCellAddress As String
'    Dim secondLanguageSubjectNameCellAddress As String
    Dim examYearClassificationCellAddress As String
    
    Dim koreanScoreCellRowNo As Integer
    Dim mathScoreCellRowNo As Integer
    
    Dim mathGradeCellAddress As String
    Dim englishGradeCellAddress As String
    Dim koreanHistoryGradeCellAddress As String
    Dim secondLanguageGradeCellAddress As String
    Dim inquiryMileStoneRowNo As Integer '탐구과목 변환표준점수 사용시 16+1, 16+2로 탐구1, 2에 대한 위치를 확인할 때, mileStoneRowNo에 16을 지정
    
    
    
'정시 sheet
    '----정시 sheet
    Dim wsJeongsi As Worksheet
    Dim jeongsiFoundRow As Long
    Dim lastRowJeongsi As Long
    Dim weightRatio(0 To 1, 0 To 27) As Variant  '반영비율과 관계된 모든 컬럼의 비율과 그 해당 값을 저장
'    Dim addBonusData() As Variant   '가산점에 대한 정보를 저장
    Dim weightRatioSKK(0, 0 To 27) As Variant  ' 성균관대학교 정보 저장
    
    Dim weightRatioColCount As Double '반영비가 저장되는 컬럼의 총 개수
    
    
    Dim englishColumn As String '영어 점수 1등급 컬럼이름
'    Dim subjectDesignatedColumn As String
    Dim jeongsiScoreProcessType As String
    Dim jeongsiScoreProcessTypeColumn As String
    Dim inquiryConvertScoreClassificationCellAddress As String
    
    
    
    '목표대학의 정보의 시작과 끝 행 번호
    Dim targetUniversityStartRow As Long ' 첫 번째 셀의 행 번호
    Dim targetUniversityEndRow As Long  ' 마지막 셀의 행 번호
    
'----정시 시트 주요 열 번호/이름 -----
    Dim colJeongsiMaxScore As String          ' 만점 열 (K열)
    Dim colJeongsiScore50Pct As String        ' 합격생 50% 점수 열 (I열)
    Dim colJeongsiScore70Pct As String        ' 합격생 70% 점수 열 (J열)
    Dim colJeongsiEnglishGrade1 As String     ' 영어 만점 열 (BK열)
    Dim colJeongsiScoreProcessType As String  ' 이미 있음 (AC열) → jeongsiScoreProcessTypeColumn과 통합 가능
    Dim colJeongsiInquiryConvert As String    ' 변환점수적용구분 열 (AD열)



'제2회국어 message를 한 번만 보여주기 위한 flag
    Dim secondLanguageFlag As Integer


'점수 계산
    Dim totalSum As Double
    

    
'가산점 정보가 들어가 있는 col의 첫번째, 마지막 번째, 컬럼의 갯수
    Dim bonusStartColNo As Long
    Dim bonusEndColNo As Long
    Dim bonusColCount As Long
    
    
'가산점에서 비율이 아니라 그냥 숫자를 넣는 변수
    Dim bonusCase1 As Double
    Dim bonusCase2 As Double
    Dim bonusCase3 As Double
    Dim bonusCase4 As Double
    Dim bonusCase7 As Double
    Dim bonusCase8 As Double
    Dim bonusCase9 As Double
    Dim bonusCase10 As Double
    
    '가산점에서 구분자가 필요
    Dim bonusFlag As Long    '성신여대 if절 구분
    
    
    Dim weightRatioStartColNo As Long     ' 검색을 시작할 열 번호 (AD열은 30번)
    Dim weightRatioEndColNo As Long       ' 검색을 종료할 열 번호 (20개 열을 검사)
    
    
    
    
    Dim wsDB_CSAT As Worksheet ' "정시 수능등급기준 수능 등급 기준
    Dim rowDB_CSATKoreanMath(1 To 2) As Integer
    Dim rowDB_CSATInquriry(1 To 2) As Integer
    
    Dim wsDB_Mock As Worksheet ' 모의고사학생입력성적유효범위 모의고사 등급 기준
    Dim rowDB_MockKoreanMath(1 To 2) As Integer
    Dim rowDB_MockInquiry(1 To 2) As Integer
    
    
    
    '합격가능성 계산을 위한 변수 지정
    Dim universityMaxScore As Double '대학발표 환산점 만점
    Dim universityScore50Pct As Double '대학발표 합격생 50%(100명 중 50등) 점수
    Dim universityScore70Pct As Double  '대학발표 합격생 70%(100명 중 70등) 점수
    Dim universityMyScore As Double '지원자의 대학 환산점수
    Dim universityPossibility As Double '지원자의 합격률 계산결과
    Dim university5070Difference As Double
    
    
' 변환표준점수 sheet 지정
    Dim wsConvertDB As Worksheet
    
    
'대학별 검색 sheet
    Dim resultDisplayRowNo As Integer
    
'대학별 검색인지, 대학별 전공별 상세검색인지 구별하기 위함.
    Dim worksheetName As String
    
    
'대학별 환산 점수를 가지고 합격 가능성 %에 대한 설명
    Dim universityPossibilityExplain As String

'표준점수 최고점을 저장하는 어레이
    Dim subjectSatStdScoreMax(0 To 4) As Double
'    Dim bSubjectSatStdScoreMaxFlag As Integer
'    Dim wsSatStdScoreMax As Worksheet
    
'대학교명, 대학교학과명

    Dim wsTargetUniversityName As Worksheet




    
Sub Main1()

    ResultCellClear
    WorksheetSetting
    ValidateKoreanMathScores
    ValidateInquiryScores
    FindTargetUniversityStartEndRows
    GetJeongsiRowNoAndProcessType
    CalculateAndDisplay             ' ← 10개 호출이 한 줄로

    
    
'    Debug.Print "국어 점수:" & koreanScore
'    Debug.Print "수학 점수:" & mathScore
'    Debug.Print "영어 점수:" & englishScore
'    Debug.Print "한국사 점수   :" & koreanHistoryScore
'    Debug.Print "탐구1 점수:" & inquiry1Score & inquiry1SubjectName
'    Debug.Print "탐구2 점수:" & inquiry2Score & inquiry2SubjectName
'    Debug.Print "제2외국어 점수:" & secondLanguageScore


    
    
End Sub

Sub Main2()
    
    Dim lastRowNo As Integer
    Dim targetUniversityGroup() As Variant
    
    ResultCellClear
    WorksheetSetting
    ValidateKoreanMathScores '국어수학 점수검증
    ValidateInquiryScores

        
    lastRowNo = wsTargetUniversityName.Cells(Rows.Count, "A").End(xlUp).Row
    
    
    resultDisplayRowNo = 17
    
    If targetUniversity = "SKY" Or _
       targetUniversity = "서성한" Or _
       targetUniversity = "중경외시이" Or _
       targetUniversity = "건동홍아숙" Or _
       targetUniversity = "국숭세단인" Or _
       targetUniversity = "광명상가" Or _
       targetUniversity = "지거국" Or _
       targetUniversity = "모든대학" Then
       
        If targetUniversity = "SKY" Then
        
            ReDim targetUniversityGroup(1 To 3)
        
            targetUniversityGroup(1) = "서울대학교"
            targetUniversityGroup(2) = "고려대학교"
            targetUniversityGroup(3) = "연세대학교"
            
        ElseIf targetUniversity = "서성한" Then
        
            ReDim targetUniversityGroup(1 To 3)
            
            targetUniversityGroup(1) = "서강대학교"
            targetUniversityGroup(2) = "성균관대학교"
            targetUniversityGroup(3) = "한양대학교"
            
        ElseIf targetUniversity = "중경외시이" Then
        
            ReDim targetUniversityGroup(1 To 5)
            
            targetUniversityGroup(1) = "중앙대학교"
            targetUniversityGroup(2) = "경희대학교"
            targetUniversityGroup(3) = "한국외국어대학교"
            targetUniversityGroup(4) = "서울시립대학교"
            targetUniversityGroup(5) = "이화여자대학교"
        
        ElseIf targetUniversity = "건동홍아숙" Then
        
            ReDim targetUniversityGroup(1 To 5)
            
            targetUniversityGroup(1) = "건국대학교"
            targetUniversityGroup(2) = "동국대학교"
            targetUniversityGroup(3) = "홍익대학교"
            targetUniversityGroup(4) = "아주대학교"
            targetUniversityGroup(5) = "숙명여자대학교"
            
            
        ElseIf targetUniversity = "국숭세단인" Then
        
            ReDim targetUniversityGroup(1 To 5)
            
            targetUniversityGroup(1) = "국민대학교"
            targetUniversityGroup(2) = "숭실대학교"
            targetUniversityGroup(3) = "세종대학교"
            targetUniversityGroup(4) = "단국대학교_죽전"
            targetUniversityGroup(5) = "인천대학교"
            
        ElseIf targetUniversity = "광명상가" Then
        
            ReDim targetUniversityGroup(1 To 4)
            
            targetUniversityGroup(1) = "광운대학교"
            targetUniversityGroup(2) = "명지대학교"
            targetUniversityGroup(3) = "상명대학교"
            targetUniversityGroup(4) = "가천대학교"
            
        ElseIf targetUniversity = "지거국" Then
        
            ReDim targetUniversityGroup(1 To 8)
            
            targetUniversityGroup(1) = "충남대학교"
            targetUniversityGroup(2) = "충북대학교"
            targetUniversityGroup(3) = "전남대학교"
            targetUniversityGroup(4) = "전북대학교"
            targetUniversityGroup(5) = "강원대학교_춘천"
            targetUniversityGroup(6) = "부산대학교"
            targetUniversityGroup(7) = "경북대학교"
            targetUniversityGroup(8) = "경상국립대학교"

        
        ElseIf targetUniversity = "모든대학" Then
            
            targetUniversityGroup = wsTargetUniversityName.Range("A10:A" & lastRowNo).Value
        End If
        
        If targetUniversity = "모든대학" Then
        
            For i = 1 To UBound(targetUniversityGroup)
                    
                targetUniversity = targetUniversityGroup(i, 1)
                FindTargetUniversityStartEndRows
                SearchTargetUniversity
            
            
            Next i
        Else
        
            For i = 1 To UBound(targetUniversityGroup)
                    
                targetUniversity = targetUniversityGroup(i)
                FindTargetUniversityStartEndRows
                SearchTargetUniversity
            
            
            Next i
        
        End If
    
    Else
    
        FindTargetUniversityStartEndRows
        SearchTargetUniversity
    End If
    
    ShowBorderLines
End Sub

Sub CalculateAndDisplay()

    InitializeScoreVariables
    ConvertTo2025CSATScore_Proportional
    GetEnglishAdmissionScore
    GetKoreanHistoryAdmissionScore
    ProcessInquiryScores
    SecondLanguage
    BonusProcess
    WeightRatioProcess
    TargetUniversityScore
    UniversityPossibilityProcess
    DisplayConvertScoreResult

End Sub




Sub InitializeScoreVariables()

    '===========================================================
    ' 국어
    ' 현재: ConvertTo2025CSATScore_Proportional 안에서 초기화
    '===========================================================
    koreanScore = 0
    koreanScoreDisplay = 0
    koreanScoreProcessType = ""

    '===========================================================
    ' 수학
    ' 현재: ConvertTo2025CSATScore_Proportional 안에서 초기화
    '===========================================================
    mathScore = 0
    mathScoreDisplay = 0
    mathScoreProcessType = ""

    '===========================================================
    ' 영어
    ' 현재: GetEnglishAdmissionScore 안에서 초기화
    '===========================================================
    englishScore = 0
    englishScoreDisplay = 0
    englishScoreProcessType = ""

    '===========================================================
    ' 한국사
    ' 현재: GetKoreanHistoryAdmissionScore 안에서 초기화
    '===========================================================
    koreanHistoryScore = 0
    koreanHistoryScoreDisplay = 0
    koreanHistoryScoreProcessType = ""

    '===========================================================
    ' 탐구1, 탐구2
    ' 현재: ProcessInquiryScores 안에서 초기화
    '===========================================================
    inquiry1SubjectName = ""
    inquiry1Score = 0
    inquiry1ScoreDisplay = 0
    inquiry1ScoreProcessType = ""

    inquiry2SubjectName = ""
    inquiry2Score = 0
    inquiry2ScoreDisplay = 0
    inquiry2ScoreProcessType = ""

    '===========================================================
    ' 제2외국어
    ' 현재: SecondLanguage 안에서 별도 초기화 없이 사용
    '       → 루프 시 이전 대학값 잔류 위험 있음
    '===========================================================
    secondLanguageScore = 0
    secondLanguageScoreDisplay = 0
    secondLanguageScoreProcessType = ""

    '===========================================================
    ' 합계 점수
    ' 현재: TargetUniversityScore의 각 대학별 Sub 시작 전
    '       totalSum = 0 이 UniversityPossibilityProcess에 있음
    '===========================================================
    totalSum = 0

    '===========================================================
    ' 합격 가능성 관련
    ' 현재: UniversityPossibilityProcess 안에서 초기화
    '===========================================================
    universityMaxScore = 0
    universityScore50Pct = 0
    universityScore70Pct = 0
    universityMyScore = 0
    universityPossibility = 0
    university5070Difference = 0
    universityPossibilityExplain = ""

    '===========================================================
    ' 가산점 관련
    ' 현재: BonusProcess 루프 안에서 매 Case마다 초기화
    '       → 루프 밖 초기화가 없어 잔류 위험
    '===========================================================
    bonusCase1 = 0
    bonusCase2 = 0
    bonusCase3 = 0
    bonusCase4 = 0
    bonusCase7 = 0
    bonusCase8 = 0
    bonusCase9 = 0
    bonusCase10 = 0


End Sub


    
Function FindMockDBRow(subjectName As String, grade As Integer) As Long
    '================================================
    ' wsDB_Mock에서 해당 과목, 등급에 맞는 행 번호를 반환
    ' 찾지 못하면 0을 반환
    '================================================
    Dim searchString As String
    Dim foundRange As Range
    
    searchString = satYearVal & satGradeVal & satMonthVal & subjectName & grade
    
    Set foundRange = wsDB_Mock.Columns(8).Find(What:=searchString, _
                                               LookIn:=xlValues, _
                                               LookAt:=xlWhole, _
                                               SearchOrder:=xlByRows, _
                                               SearchDirection:=xlNext, _
                                               MatchCase:=False)
    
    If Not foundRange Is Nothing Then
        FindMockDBRow = foundRange.Row
    Else
        FindMockDBRow = 0
    End If

End Function




Sub SearchTargetUniversity()

    Dim i As Integer

    For i = targetUniversityStartRow To targetUniversityEndRow
        If wsJeongsi.Cells(i, "B").Value = targetUniversity Then

            targetMajor = wsJeongsi.Cells(i, "E").Value
            applicationType = wsJeongsi.Cells(i, "D").Value
            jeongsiFoundRow = i
            jeongsiScoreProcessType = wsJeongsi.Cells(jeongsiFoundRow, jeongsiScoreProcessTypeColumn).Value

            CalculateAndDisplay             ' ← InitializeScoreVariables 포함
            resultDisplayRowNo = resultDisplayRowNo + 1

        End If
    Next i
'            Debug.Print "정시 행번호: "; jeongsiFoundRow
'            Debug.Print "국어 점수:" & koreanScore
'            Debug.Print "수학 점수:" & mathScore
'            Debug.Print "영어 점수:" & englishScore
'            Debug.Print "한국사 점수   :" & koreanHistoryScore
'            Debug.Print "탐구1 점수:" & inquiry1Score
'            Debug.Print "탐구2 점수:" & inquiry2Score
'            Debug.Print "제2외국어 점수:" & secondLanguageScore
'            Debug.Print "합격 능서:" & universityPossibilityExplain
End Sub


Sub WorksheetSetting()
    Dim yearPrefix As String
    
    
    worksheetName = ThisWorkbook.ActiveSheet.Name
    
    '============================================================
    '대학 학과별 상세검색 이름을 대학,학과별 상세 검색으로 바꿀 예정
    '대학별 검색 sheet를 추가로 만들 예정
    '============================================================
    If worksheetName = "대학 학과별 상세검색" Then
        Set wsStudentData = ThisWorkbook.Sheets("대학 학과별 상세검색")  '대학 학과별 상세검색
    Else
        Set wsStudentData = ThisWorkbook.Sheets("대학별검색")
    End If
    
    Set wsDB_CSAT = ThisWorkbook.Sheets("정시 수능등급기준") ' 수능 등급 기준
    Set wsDB_Mock = ThisWorkbook.Sheets("모의고사학생입력성적유효범위") ' 수능/모의고사 학생성적 입력 점수 국어, 수학 등급별 유효성 기준
    
    englishColumn = "BK"
    jeongsiScoreProcessTypeColumn = "AC"
'    subjectDesignatedColumn = "DB"
    jeongsiScoreProcessType = ""
    applicationType = ""
    inquiryConvertScoreClassificationCellAddress = "AD" '변환점수적용구분
    
    bonusStartColNo = 90 'CL열 가산점의 시작 컬럼 번호
    bonusEndColNo = 103    'CY열 가산점의 끝 컬럼 번호
    bonusColCount = bonusEndColNo - bonusStartColNo + 1 '보너스 정보가 저장된 컬럼의 총개수 bonusEndColNo(102)- bonusStartColNo(89)+1=14
    
    ' --- [설정 부분] ---
    ' weightRatioStartColNo과 weightRatioEndColNo은 이 프로시저 내에서 정의합니다.
    
    weightRatioStartColNo = 31       ' AE열 번호 31
    weightRatioEndColNo = 57         ' BE열 번호 57
    weightRatioColCount = weightRatioEndColNo - weightRatioStartColNo + 1   '반영비율이 저장되는 열의 갯수 현재는 27개 AE(31)-BE(57)
    
    secondLanguageFlag = 0
    
    ' --- (설정 끝) ---
    
    
    satYearValCellAddress = "C2"    '년도
    satGradeValCellAddress = "C3" '학년
    satMonthValCellAddress = "C4"   '시행 월

    If worksheetName = "대학별검색" Then
    
        satYearValCellAddress = "C3"    '년도
        satGradeValCellAddress = "D3" '학년
        satMonthValCellAddress = "E3"   '시행 월
    
    End If

    If worksheetName = "대학 학과별 상세검색" Then
        targetUniversityCellAddress = "C7"
        targetMajorCellAddress = "C8"
        koreanCellAddress = "A13"
        koreanSubjectNameCellAddress = "B13"
        
        mathCellAddress = "A14"
        mathSubjectNameCellAddress = "B14"
        
        
        examYearClassificationCellAddress = "F12"
        
        koreanScoreCellRowNo = 13
        mathScoreCellRowNo = 14
        
        mathGradeCellAddress = "E14"
        englishGradeCellAddress = "E15"
        koreanHistoryGradeCellAddress = "E16"
        secondLanguageGradeCellAddress = "E19"
        inquiryMileStoneRowNo = 16
        

    ElseIf worksheetName = "대학별검색" Then

        targetUniversityCellAddress = "B14"
'        targetMajorCellAddress = "C8"
        koreanCellAddress = "A5"
        koreanSubjectNameCellAddress = "B5"
        
        mathCellAddress = "A6"
        mathSubjectNameCellAddress = "B6"
        
        
        examYearClassificationCellAddress = "B13"

        koreanScoreCellRowNo = 5
        mathScoreCellRowNo = 6
        
        mathGradeCellAddress = "E6"
        englishGradeCellAddress = "E7"
        koreanHistoryGradeCellAddress = "E8"
        secondLanguageGradeCellAddress = "E11"
        inquiryMileStoneRowNo = 8



    End If
    
    
    
    
    
    '================================================
    ' 대학 학과별 상세검색 시트에서 변환 기준 수능, 대학명과 학과명 가져오기
    ' 예) 2025학년도 수능
    '================================================
    
    targetUniversity = wsStudentData.Range(targetUniversityCellAddress).Value
    
    If worksheetName = "대학별검색" Then
        targetMajor = ""
    Else
    
        targetMajor = Replace(wsStudentData.Range(targetMajorCellAddress).Value, "*", "~*")
    End If
    
    examYearClassification = wsStudentData.Range(examYearClassificationCellAddress).Value
    
    
    
'    If examYearClassification = "2024학년도 정시" Then
'
'        Set wsJeongsi = ThisWorkbook.Sheets("24정시")
'        Set wsConvertDB = ThisWorkbook.Sheets("24정시 변환표준점수") ' 변환표 DB 시트 이름
'        Set wsTargetUniversityName = ThisWorkbook.Sheets("24정시 대학명 학과명")
'
'    ElseIf examYearClassification = "2025학년도 정시" Then
'
'        Set wsJeongsi = ThisWorkbook.Sheets("25정시")
'        Set wsConvertDB = ThisWorkbook.Sheets("25정시 변환표준점수") ' 변환표 DB 시트 이름
'        Set wsTargetUniversityName = ThisWorkbook.Sheets("25정시 대학명 학과명")
'
'    End If
    '=======================================================================
    '윗 부분을 아래로 수정한 것임
    '=======================================================================
    
    yearPrefix = Mid(examYearClassification, 3, 2)  ' "2024학년도 정시" → "24"
    Set wsJeongsi = ThisWorkbook.Sheets(yearPrefix & "정시")
    Set wsConvertDB = ThisWorkbook.Sheets(yearPrefix & "정시 변환표준점수")
    Set wsTargetUniversityName = ThisWorkbook.Sheets(yearPrefix & "정시 대학명 학과명")
'    Set wsSatStdScoreMax = ThisWorkbook.Sheets("수능표준점수최고점")
    
'    For i = LBound(subjectSatStdScoreMax) To UBound(subjectSatStdScoreMax)
'
'        subjectSatStdScoreMax(i) = 0
'    Next i
    
'    bSubjectSatStdScoreMaxFlag = 0
    
    ' WorksheetSetting에 추가
    colJeongsiMaxScore = "K"         ' 만점
    colJeongsiScore50Pct = "I"       ' 합격생 50% 점수
    colJeongsiScore70Pct = "J"       ' 합격생 70% 점수
    colJeongsiEnglishGrade1 = "BK"  ' 영어 만점


    
End Sub


Sub FindTargetUniversityStartEndRows()
 
    Dim searchRange As Range
    Dim firstFound As Range ' 첫 번째 찾은 셀 객체
    Dim lastFound As Range  ' 마지막 찾은 셀 객체
'    Dim targetChar As String
    Dim colNum As Long
    

    
    ' 검색할 열 번호 설정 (예: A열은 1, B열은 2 등)
    colNum = 2 ' <--- 여기를 원하는 열 번호로 변경하세요 (현재는 B열)


    'targetUniversityStartRow, targetUniversityEndRow 초기화
    targetUniversityStartRow = 0
    targetUniversitySendRow = 0
    
    ' 해당 열 전체를 검색 범위로 설정
    Set searchRange = wsJeongsi.Columns(colNum)

    ' 1. 첫 번째 찾기 (전방 검색: 위에서 아래로)
    '찾을 문자열 targetUniversity
    Set firstFound = searchRange.Find(What:=targetUniversity, _
                                      LookIn:=xlValues, _
                                      LookAt:=xlWhole, _
                                      SearchOrder:=xlByRows, _
                                      SearchDirection:=xlNext, _
                                      MatchCase:=False)

    ' 2. 마지막 찾기 (후방 검색: 아래에서 위로)
    ' (주의: firstFound가 Nothing이 아니어야 lastFound를 찾는 의미가 있어요)
    If Not firstFound Is Nothing Then
        Set lastFound = searchRange.Find(What:=targetUniversity, _
                                        LookIn:=xlValues, _
                                        LookAt:=xlWhole, _
                                        SearchOrder:=xlByRows, _
                                        SearchDirection:=xlPrevious, _
                                        MatchCase:=False)
                                        
        ' 찾은 셀이 있다면 해당 행 번호를 변수에 저장
        targetUniversityStartRow = firstFound.Row
        targetUniversityEndRow = lastFound.Row
        
        ' 결과를 한 번에 출력
'        MsgBox "'" & targetChar & "' 데이터 블록 정보:" & vbCrLf & _
'               "첫 행 번호 (targetUniversityStartRow): " & targetUniversityStartRow & "행" & vbCrLf & _
'               "마지막 행 번호 (targetUniversityEndRow): " & targetUniversityEndRow & "행"
    Else
        MsgBox "'" & targetUniversity & "' 데이터를 찾을 수 없습니다."
    End If

End Sub

Sub ValidateKoreanMathScores()

    '국어수학 점수검증 모듈
    
    Dim subjectsName(1 To 2) As String ' 과목명 배열
    Dim rowsToCheck(1 To 2) As Long ' 각 과목의 행 번호 배열
    Dim i As Integer, lastRowDB As Long, j As Long
    Dim rankVal As Integer, scoreVal As Double
    Dim minScore As Double, maxScore As Double
    Dim matchFound As Boolean
    Dim alertMsg As String
    Dim errorCount As Integer
    
    
    Dim searchMockString As String
    Dim foundMockRange As Range
    Dim searchMockRange As Range
    Dim foundMockRow As String
    Dim foundCSAT As Boolean, foundMock As Boolean
        
    ' 대학 학과별 상세검색 시트의 고정된 위치에서 년도, 학년, 회차/월 값 가져오기
    satYearVal = wsStudentData.Range(satYearValCellAddress).Value ' 예: C2에 년도(2025)
    satGradeVal = wsStudentData.Range(satGradeValCellAddress).Value ' 예: C3에 학년(고3)
    satMonthVal = wsStudentData.Range(satMonthValCellAddress).Value ' 예: C4에 회차/월(9월)
    foundMock = False
    
    For i = 1 To 2 ' 첫 번째 차원 (0 또는 1)
        subjectsName(i) = Empty
        rowsToCheck(i) = 0
        rowDB_MockKoreanMath(i) = 0
    Next i
       
    ' 검증할 과목명과 해당 행 번호 설정
    subjectsName(1) = wsStudentData.Range(koreanCellAddress).Value ' 국어 과목명 (A13 셀), 상세과목명 아님
    rowsToCheck(1) = koreanScoreCellRowNo ' 국어 성적 입력 행
    
    subjectsName(2) = wsStudentData.Range(mathCellAddress).Value ' 수학 과목명 (A14 셀), 상세과목명 아님
    rowsToCheck(2) = mathScoreCellRowNo ' 수학 성적 입력 행
    
    ' 모의고사학생입력성적유효범위 시트의 마지막 데이터 행 찾기
    lastRowDB = wsDB_Mock.Cells(wsDB_Mock.Rows.Count, "A").End(xlUp).Row
    
    alertMsg = "검증 결과 - 표준점수 범위 벗어난 과목:" & vbNewLine
    errorCount = 0
    
    ' 국어와 수학 두 과목에 대해 반복하여 검증
    
    Set searchMockRange = wsDB_Mock.Columns(8)
    
    For i = 1 To 2
        ' 현재 과목의 등급(E열)과 표준점수(C열) 값 가져오기
        rankVal = wsStudentData.Cells(rowsToCheck(i), "E").Value ' 등급은 E열
        scoreVal = wsStudentData.Cells(rowsToCheck(i), "C").Value ' 표준점수는 C열
        
        matchFound = False
        
        ' DB 시트에서 해당 과목, 년도, 학년, 월, 등급에 맞는 표준점수 범위 찾기
'        For j = 2 To lastRowDB ' DB 시트의 2행부터 마지막 행까지
'            If wsDB_Mock.Cells(j, "A").Value = satYearVal And _
'               wsDB_Mock.Cells(j, "B").Value = satGradeVal And _
'               wsDB_Mock.Cells(j, "C").Value = satMonthVal And _
'               wsDB_Mock.Cells(j, "D").Value = subjectsName(i) And _
'               wsDB_Mock.Cells(j, "E").Value = rankVal Then
'
'               minScore = wsDB_Mock.Cells(j, "F").Value ' DB 시트 F열 (표준점수 최소값)
'               maxScore = wsDB_Mock.Cells(j, "G").Value ' DB 시트 G열 (표준점수 최대값)
'               matchFound = True
'               Exit For ' 일치하는 데이터를 찾았으므로 더 이상 검색할 필요 없음
'            End If
'        Next j

        rowDB_MockKoreanMath(i) = FindMockDBRow(subjectsName(i), rankVal)
        If rowDB_MockKoreanMath(i) = 0 Then
            MsgBox subjectsName(i) & " 과목의 DB 데이터를 찾을 수 없습니다.", vbExclamation
            Exit Sub
        End If
        
        minMockScore = wsDB_Mock.Cells(rowDB_MockKoreanMath(i), "F").Value
        maxMockScore = wsDB_Mock.Cells(rowDB_MockKoreanMath(i), "G").Value
        
        If scoreVal < minMockScore Or scoreVal > maxMockScore Then
            alertMsg = alertMsg & subjectsName(i) & " - 등급 " & rankVal & _
                       ", 표준점수 " & scoreVal & _
                       " (기준 " & minMockScore & " ~ " & maxMockScore & ")" & vbNewLine
            errorCount = errorCount + 1
        End If
    Next i
    
    ' 최종 검증 결과 요약 메시지 출력
    If errorCount > 0 Then
        MsgBox alertMsg, vbExclamation, "성적 검증 오류"
    Else
'        MsgBox "국어와 수학 표준점수가 모두 적정 범위 내에 있습니다.", vbInformation, "성적 검증 완료"
    End If
    
End Sub


Sub ValidateInquiryScores()
    ' 변수 선언 (Dim)
    Dim inquirySubjectName(1 To 2) As String
    Dim inquiryStdScore(1 To 2) As Double
    Dim inquiryPercentile(1 To 2) As Double
    Dim inquiryGrade(1 To 2) As Integer
    Dim inquiryRows(1 To 2) As Integer
'    Dim convertScoreResult(1 To 2) As Variant '국어, 수학 변호나 점수 결과 저장
'    Dim convertScoreTypeResult(1 To 2) As String '대학 국어, 수학 반영 기준 예) 백분위, 표준점수, 등급, 변환표준점수
    
'    Dim outputResult(1 To 2) As Variant ' 최종 결과 (점수 또는 텍스트)
'    Dim outputCells(1 To 2) As Range   ' 결과를 출력할 셀 (F17, F18)
    Dim i As Integer ' 탐구 과목 루프 카운터
'    Dim roundedPercentile As Integer
    
    
'    Dim convertDBUnivCol As Long ' 대학별 탐구 변환표준점수표 DB 시트의 대학 컬럼
'    Dim convertDBPercentileRow As Long
'    Dim lastRowJeongsi As Long, lastColJeongSi As Long
'    Dim lastRowConvertDB As Long, lastColConvertDB As Long
    
'    Dim lookupUnivName As String ' 변환표준점수 DB에서 대학 컬럼을 찾을 때 사용할 이름 (고려대-과탐/사탐 등)
'    Dim inquiryConvertScoreClassificationType As String
    
    ' For 루프 내에서 사용될 변수들은 미리 선언
'    Dim col As Long
'    Dim rowCSAT As Long, rowMock As Long
'    Dim lastChar As String
    Dim satExamYearSubjectGradeString As String
    Dim satExamYearSubjectGradeFoundRange As Range
    Dim satExamYearSubjectGradeRange As Range
    Dim satExamYearSubjectGradeRow As String
    
    
'    Dim searchMockString As String
'    Dim foundMockRange As Range
'    Dim searchMockRange As Range
'    Dim foundMockRow As String
    
    
    
    
    '해당 시험이 표준점수를 비교대상 정시의 표준점수로 변환하기 위한 변수정의
'    Dim inputStdScore As Double
'    Dim inputPercentile As Double
'    Dim inputGrade As Integer
    Dim minMockScore As Double, maxMockScore As Double
    Dim minCSATScore As Double, maxCSATScore As Double
'    Dim relativePos As Double
'    Dim convertedScore As Double
    Dim foundCSAT As Boolean, foundMock As Boolean
    Dim lastRowCSAT As Long, lastRowMock As Long
    
    Dim alertMsg As String
    Dim validateErrorFlag As Integer
    
    ' 1. 시트 설정
    
    ' 2. '대학 학과별 상세검색' 시트에서 공통 정보 가져오기
    '이 프로시저에서 필요한 배열의 초기화
     For i = 1 To 2 ' 첫 번째 차원 (0 또는 1)
        inquirySubjectName(i) = Empty
        inquiryStdScore(i) = 0
        inquiryPercentile(i) = 0
        inquiryGrade(i) = 0
'        Set outputCells(i) = Nothing
'        convertScoreResult(i) = 0
'        convertScoreTypeResult(i) = Empty
        inquiryRows(i) = 0
        rowDB_MockInquiry(i) = 0
     Next i
     
     
    lastChar = ""
    alertMsg = "검증 결과 - 표준점수 범위 벗어난 과목:" & vbNewLine
    validateErrorFlag = 0
    Set searchMockRange = wsDB_Mock.Columns(8)
     
    lastRowCSAT = wsDB_CSAT.Cells(wsDB_CSAT.Rows.Count, "A").End(xlUp).Row
    lastRowMock = wsDB_Mock.Cells(wsDB_Mock.Rows.Count, "A").End(xlUp).Row
     
    ' --- (초기화 끝) ---
    
    
    ' 6. 탐구 1, 2 과목별로 점수 처리
    For i = 1 To 2 ' 1은 탐구1 (17행), 2는 탐구2 (18행)
        ' '대학 학과별 상세검색' 시트에서 각 탐구 과목 정보 읽어오기
        inquirySubjectName(i) = Trim(wsStudentData.Cells(inquiryMileStoneRowNo + i, "B").Value) ' B17, B18 (과목명), Trim으로 공백 제거
        
        ' 탐구 과목이 입력되었는지 확인
        If inquirySubjectName(i) = "" Then
'            outputCells(i).Value = "과목없음" ' 탐구 과목이 없으면 '과목없음'으로 표시
            GoTo NextSubject ' 다음 탐구 과목으로 바로 넘어감
            
        '2022개정 교육과정의 학생이 통합사회, 통합과학으로 과목을 선택하였을 때, 이전 학년도의 수능 선택과목으로 강제 치환
'        ElseIf inquirySubjectName(i) = "사회탐구" And _
'                (examYearClassification = "2024학년도 정시" Or _
'                examYearClassification = "2025학년도 정시" Or _
'                 examexamYearClassification = "2026학년도 정시" Or _
'                 examYearClassification = "2027학년도 정시" _
'                 ) Then
'
'            inquirySubjectName(i) = "사회·문화"    '2022개정교육과정 학생이
'
'        ElseIf inquirySubjectName(i) = "과학탐구" And _
'                (examYearClassification = "2024학년도 정시" Or _
'                examYearClassification = "2025학년도 정시" Or _
'                 examexamYearClassification = "2026학년도 정시" Or _
'                 examYearClassification = "2027학년도 정시" _
'                 ) Then
'            inquirySubjectName(i) = "지구과학Ⅰ"
        
        End If

        inquiryStdScore(i) = wsStudentData.Cells(inquiryMileStoneRowNo + i, "C").Value
        inquiryPercentile(i) = wsStudentData.Cells(inquiryMileStoneRowNo + i, "D").Value
        inquiryGrade(i) = wsStudentData.Cells(inquiryMileStoneRowNo + i, "E").Value
        inquiryRows(i) = inquiryMileStoneRowNo + i
        
        ' 변경 후
        rowDB_MockInquiry(i) = FindMockDBRow(inquirySubjectName(i), inquiryGrade(i))
        If rowDB_MockInquiry(i) = 0 Then
            MsgBox inquirySubjectName(i) & " 과목의 DB 데이터를 찾을 수 없습니다.", vbExclamation
            Exit Sub
        End If
        
        minMockScore = wsDB_Mock.Cells(rowDB_MockInquiry(i), "F").Value
        maxMockScore = wsDB_Mock.Cells(rowDB_MockInquiry(i), "G").Value
        
        If inquiryStdScore(i) < minMockScore Or inquiryStdScore(i) > maxMockScore Then
            alertMsg = alertMsg & inquirySubjectName(i) & " - 등급 " & inquiryGrade(i) & _
                       ", 표준점수 " & inquiryStdScore(i) & _
                       " (기준 " & minMockScore & " ~ " & maxMockScore & ")" & vbNewLine
            validateErrorFlag = validateErrorFlag + 1
        End If
NextSubject: ' GoTo 문을 위한 레이블
    Next i
    

    If validateErrorFlag > 0 Then
    
        MsgBox alertMsg, vbExclamation, "성적 검증 오류"
    End If
    


End Sub



Sub GetJeongsiRowNoAndProcessType()

    Dim lastRowJeongsi As Long
    
    ' 정시 시트에서 해당 대학, 학과에 맞는 행 찾기
    lastRowJeongsi = 0
    lastRowJeongsi = wsJeongsi.Cells(Rows.Count, "B").End(xlUp).Row
    


    '================================================
    '검색 대상 기준 수능 예) 2025학년도 정시
    '목표대학교, 목표학과의 해당 행 번호를 찾는다
    '1행은 제목 행이라 2행부터 시작
    '위에서 찾은 해당 대학의 시작행에서 끝행까지
    '================================================
    jeongsiFoundRow = 0
    
    For i = targetUniversityStartRow To targetUniversityEndRow
    
        If applicationType = "" Then
            If wsJeongsi.Cells(i, "A").Value = examYearClassification And _
                wsJeongsi.Cells(i, "B").Value = targetUniversity And _
                wsJeongsi.Cells(i, "E").Value = targetMajor Then
                jeongsiFoundRow = i
                Exit For
            End If
        Else
            If wsJeongsi.Cells(i, "A").Value = examYearClassification And _
                wsJeongsi.Cells(i, "B").Value = targetUniversity And _
                wsJeongsi.Cells(i, "E").Value = targetMajor And _
                wsJeongsi.Cells(i, "D").Value = applicationType Then
                jeongsiFoundRow = i
                Exit For
            End If
        
        End If
    Next i
    
    If jeongsiFoundRow = 0 Then
        MsgBox "정시 시트에서 해당 대학, 학과 정보가 없습니다. 점수 변환을 수행할 수 없습니다.", vbExclamation
        ' 모든 F열과 G열 초기화 및 오류 메시지 출력
'        For i = 13 To 19
'            wsStudentData.Cells(i, "F").Value = "정보없음"
'            wsStudentData.Cells(i, "G").Value = "" ' G열은 비워둠 (요청 반영)
'        Next i
        Exit Sub
    End If
    
    ' ??? 가장 중요한 수정 부분: AC열의 점수 처리 방식을 딱 한 번만 가져와서 변수에 저장 ???
    jeongsiScoreProcessType = wsJeongsi.Cells(jeongsiFoundRow, jeongsiScoreProcessTypeColumn).Value



End Sub

Sub ConvertTo2025CSATScore_Proportional()
    Dim lastRowCSAT As Long, lastRowMock As Long
    Dim subjectsName(1 To 2) As String ' 국어, 수학 과목명
    Dim inputRows(1 To 2) As Long ' 국어, 수학 행
    Dim convertScoreResult(1 To 2) As Double '국어, 수학 변호나 점수 결과 저장
    Dim convertScoreTypeResult(1 To 2) As String '대학 국어, 수학 반영 기준 예) 백분위, 표준점수, 등급, 변환표준점수
    Dim i As Integer, rowCSAT As Long, rowMock As Long
    Dim inputStdScore As Double
    Dim inputPercentile As Double
    Dim inputGrade As Integer
    Dim minMockScore As Double, maxMockScore As Double
    Dim minCSATScore As Double, maxCSATScore As Double
    Dim relativePos As Double
    Dim convertedScore As Double
    Dim foundCSAT As Boolean, foundMock As Boolean
    
    
    Dim searchSatString As String
    Dim foundSatRange As Range
    Dim searchSatRange As Range
    Dim foundSatRow As String
    Dim validateErrorFlag As Integer
    
'    Dim jeongSiScoreProcessType As String ' AC열에서 가져올 통합 처리 유형
    
    
'    ' 정시 시트에서 해당 대학, 학과에 맞는 행 찾기
'    lastRowJeongSi = wsJeongsi.Cells(Rows.count, "B").End(xlUp).Row
'    jeongSiFoundRow = 0
    ' 추가 변수 선언
    Dim satMaxString As String
    Dim foundSatMaxRange As Range
    Dim foundSatMaxRow As Long
    
    
    '이 프로시저에서 필요한 배열의 초기화
     For i = 1 To 2 ' 첫 번째 차원 (0 또는 1)
        subjectsName(i) = Empty
        inputRows(i) = 0
        convertScoreResult(i) = 0
        convertScoreTypeResult(i) = Empty
     Next i
    ' --- (초기화 끝) ---
   
    subjectsName(1) = wsStudentData.Range(koreanCellAddress).Value ' 국어,  선택 과목명 아님
    inputRows(1) = koreanScoreCellRowNo
    
    subjectsName(2) = wsStudentData.Range(mathCellAddress).Value ' 수학 과목명, 선택 과목명 아님.확률과 통계X,  미적분X, 기하X
    inputRows(2) = mathScoreCellRowNo
    
    lastRowCSAT = wsDB_CSAT.Cells(wsDB_CSAT.Rows.Count, "A").End(xlUp).Row
    lastRowMock = wsDB_Mock.Cells(wsDB_Mock.Rows.Count, "A").End(xlUp).Row
    
    
    Set searchSatRange = wsDB_CSAT.Columns(6)
    
    ' 국어(i=1), 수학(i=2) 점수 처리 (F열 계산 및 G열 표시)
    For i = 1 To 2
        inputStdScore = wsStudentData.Cells(inputRows(i), "C").Value   ' 모의고사 표준점수 (C열)
        inputPercentile = wsStudentData.Cells(inputRows(i), "D").Value ' 모의고사 백분위 (D열)
        inputGrade = wsStudentData.Cells(inputRows(i), "E").Value   ' 모의고사 등급 (E열)
        validateErrorFlag = 0
        
        Select Case jeongsiScoreProcessType ' ? AC열의 통합 처리 유형 사용 ?
            Case "표준점수", "변환표준점수"
                ' F열: 국수 점수 변환 로직 (표준점수/변환표준점수 모두 동일 로직 적용)
                foundMock = False
                foundCSAT = False
                
'                For rowMock = 2 To lastRowMock
'                    '=======================================================
'                    ' 2025
'                    ' 고3
'                    '11월 수능
'                    '국어, 수학
'                    '등급
'                    '=======================================================
'                    If wsDB_Mock.Cells(rowMock, "A").Value = satYearVal And _
'                        wsDB_Mock.Cells(rowMock, "B").Value = satGradeVal And _
'                        wsDB_Mock.Cells(rowMock, "C").Value = satMonthVal And _
'                        wsDB_Mock.Cells(rowMock, "D").Value = subjectsName(i) And _
'                        wsDB_Mock.Cells(rowMock, "E").Value = inputGrade Then
'                        minMockScore = wsDB_Mock.Cells(rowMock, "F").Value
'                        maxMockScore = wsDB_Mock.Cells(rowMock, "G").Value
'                        foundMock = True
'                        Exit For
'                    End If
'                Next rowMock
                
                If rowDB_MockKoreanMath(i) > 0 Then
                    minMockScore = wsDB_Mock.Cells(rowDB_MockKoreanMath(i), "F").Value
                    maxMockScore = wsDB_Mock.Cells(rowDB_MockKoreanMath(i), "G").Value
                    foundMock = True
                End If
                
                
                
                searchSatString = examYearClassification & subjectsName(i) & inputGrade
                Set foundSatRange = searchSatRange.Find(What:=searchSatString, _
                                                    LookIn:=xlValues, _
                                                    LookAt:=xlWhole, _
                                                    SearchOrder:=xlByRows, _
                                                    SearchDirection:=xlNext, _
                                                    MatchCase:=False)
        
                If Not foundSatRange Is Nothing Then
        
                    ' 찾은 셀이 있다면 해당 행 번호를 변수에 저장
                    foundSatRow = foundSatRange.Row
                    minCSATScore = wsDB_CSAT.Cells(foundSatRow, "D").Value
                    maxCSATScore = wsDB_CSAT.Cells(foundSatRow, "E").Value
                    foundCSAT = True
        
                Else
                    MsgBox "'" & targetChar & "' 데이터를 찾을 수 없습니다."
                    Exit Sub
                End If
                
                '========================================
                ' 추가: 과목 최고점을 1등급 고정으로 조회
                '========================================
                satMaxString = examYearClassification & subjectsName(i) & 1
                Set foundSatMaxRange = searchSatRange.Find(What:=satMaxString, _
                                                           LookIn:=xlValues, _
                                                           LookAt:=xlWhole, _
                                                           SearchOrder:=xlByRows, _
                                                           SearchDirection:=xlNext, _
                                                           MatchCase:=False)
                If Not foundSatMaxRange Is Nothing Then
                    foundSatMaxRow = foundSatMaxRange.Row
                    If i = 1 Then subjectSatStdScoreMax(0) = wsDB_CSAT.Cells(foundSatMaxRow, "E").Value  '국어 최고점
                    If i = 2 Then subjectSatStdScoreMax(1) = wsDB_CSAT.Cells(foundSatMaxRow, "E").Value  '수학 최고점
                End If
                
                '========================================
                
                
                If foundCSAT Then
                    If inputStdScore < minCSATScore Or inputStdScore > maxCSATScore Then
                        ' 입력한 표준점수가 범위를 벗어났을 경우
                        alertMsg = alertMsg & subjectsName(i) & " - 등급 " & inputGrade & ", 표준점수 " & inputStdScore & " (기준 " & minCSATScore & " ~ " & maxCSATScore & ")" & vbNewLine
                        validateErrorFlag = validateErrorFlag + 1
                    End If
                End If
                
                
                
                
'                For rowCSAT = 2 To lastRowCSAT
'                    '=======================================================
'                    '2025학년도 수능
'                    '국어, 수학
'                    '등급
'                    '=======================================================
'                    If wsDB_CSAT.Cells(rowCSAT, "A").Value = examYearClassification And _
'                        wsDB_CSAT.Cells(rowCSAT, "B").Value = subjectsName(i) And _
'                        wsDB_CSAT.Cells(rowCSAT, "C").Value = inputGrade Then
'                        minCSATScore = wsDB_CSAT.Cells(rowCSAT, "D").Value
'                        maxCSATScore = wsDB_CSAT.Cells(rowCSAT, "E").Value
'                        foundCSAT = True
'                        Exit For
'                    End If
'                Next rowCSAT
'
'
'                If foundMock And foundCSAT Then
'                    If scoreVal < minScore Or scoreVal > maxScore Then
'                        ' 입력한 표준점수가 범위를 벗어났을 경우
'                        alertMsg = alertMsg & subjectsName(i) & " - 등급 " & rankVal & ", 표준점수 " & scoreVal & " (기준 " & minScore & " ~ " & maxScore & ")" & vbNewLine
'                        errorCount = errorCount + 1
'                    End If
'                End If
                
                
                If foundMock And foundCSAT Then
                    If maxMockScore = minMockScore Then
                        relativePos = 0
                    Else
                        relativePos = (inputStdScore - minMockScore) / (maxMockScore - minMockScore)
                    End If
                    
                    If relativePos < 0 Then relativePos = 0
                    If relativePos > 1 Then relativePos = 1
                    
                    convertedScore = minCSATScore + relativePos * (maxCSATScore - minCSATScore)
                    convertScoreResult(i) = CLng(Application.WorksheetFunction.Round(convertedScore, 0))
                    convertScoreTypeResult(i) = "표준점수"
                Else
                    If worksheetName = "대학 학과별 상세검색" Then
                        wsStudentData.Cells(inputRows(i), "F").Value = "기준없음"
                    End If
                End If
                

            Case "백분위"
                convertScoreResult(i) = inputPercentile ' F열: 백분위 값을 그대로 사용
                convertScoreTypeResult(i) = "백분위" ' ? G열: "백분위" 표시 ?

            Case "등급"
                convertScoreResult(i) = inputGrade ' F열: 등급 값을 그대로 사용
                convertScoreTypeResult(i) = "등급" ' ? G열: "등급" 표시 ?
                
            Case Else
                If worksheetName = "대학 학과별 상세검색" Then
                    wsStudentData.Cells(inputRows(i), "F").Value = "처리유형오류"
                    wsStudentData.Cells(inputRows(i), "G").Value = "처리유형오류"
                End If
        End Select
    Next i
    
    '============================================================
    '국어와 수학의 변화 점수를 저장한 것은 Module단위 변수에 저장
    'DisplayCoｎvertScoreResult 포르시져에서 셀에 값을 할당한다.
    '============================================================
    koreanScore = 0
    koreanScoreDisplay = 0
    mathScore = 0
    mathScoreDisplay = 0
    koreanScoreProcessType = ""
    mathScoreProcessType = ""
    
    For i = 1 To 2
    
        If IsNumeric(convertScoreResult(i)) And convertScoreResult(i) > 0 Then
        
            If i = 1 Then
            
                koreanScore = convertScoreResult(i)
                koreanScoreDisplay = convertScoreResult(i)
                koreanScoreProcessType = convertScoreTypeResult(i)
            
            ElseIf i = 2 Then
                
                mathScore = convertScoreResult(i)
                mathScoreDisplay = convertScoreResult(i)
                mathScoreProcessType = convertScoreTypeResult(i)
            
            End If
        End If
     Next i
    
'    MsgBox "점수 처리(2025 수능 기준)가 완료되었습니다.", vbInformation
End Sub




Sub GetEnglishAdmissionScore()
    Dim englishGrade As Integer
    Dim targetColumnHeader As String
    Dim targetCol As Long
    Dim lastColJeongSi As Long
    
    englishGrade = 0
    englishGrade = wsStudentData.Range(englishGradeCellAddress).Value   ' 영어 등급 (E15 셀)
    
    ' 영어 등급이 유효한 범위(1~9)인지 확인
    If englishGrade < 1 Or englishGrade > 9 Then
        MsgBox "대학 학과별 상세검색 시트의 영어 등급(E15)이 1~9 범위를 벗어났습니다.", vbCritical
        Exit Sub
    End If
    
    
    ' 4. 영어 등급에 맞는 컬럼 헤더 찾기 (예: "영1", "영2"...)
    targetColumnHeader = "영" & englishGrade
    
    lastColJeongSi = wsJeongsi.Cells(1, Columns.Count).End(xlToLeft).Column ' 1행 헤더의 마지막 컬럼 찾기
    targetCol = 0 ' 초기화
    
    For c = 1 To lastColJeongSi
        If wsJeongsi.Cells(1, c).Value = targetColumnHeader Then ' 1행은 헤더라고 가정
            targetCol = c
            Exit For
        End If
    Next c
    
    If targetCol = 0 Then
        MsgBox "정시 시트에서 '" & targetColumnHeader & "' 컬럼을 찾을 수 없습니다.", vbExclamation
        wsStudentData.Range("F14").Value = "컬럼없음" ' 변환 결과를 넣을 셀 F14 가정
        Exit Sub
    End If
    
    ' 5. 해당하는 영어 점수 가져오기
    englishScore = 0
    englishScoreDisplay = 0
    englishScore = wsJeongsi.Cells(jeongsiFoundRow, targetCol).Value
    englishScoreDisplay = wsJeongsi.Cells(jeongsiFoundRow, targetCol).Value
    ' 6. '대학 학과별 상세검색' 시트에 결과 출력 (F15 셀에 출력한다고 가정)
'    wsStudentData.Range("F15").Value = englishScore
    '========================================
    ' 추가: 영어 최고점 저장 (1등급 변환점수)
    '========================================
    subjectSatStdScoreMax(2) = wsJeongsi.Cells(jeongsiFoundRow, colJeongsiEnglishGrade1).Value
    '========================================


    englishScoreProcessType = ""
    
    If IsNumeric(englishScore) Then
        englishScoreProcessType = "변환점수"
    Else
        englishScoreProcessType = ""
    End If
    
'    englishScore = wsStudentData.Cells(15, "F").Value
'    MsgBox "영어 반영 점수를 성공적으로 가져왔습니다: " & englishScore, vbInformation
    
End Sub


Sub GetKoreanHistoryAdmissionScore()
    Dim koreanHistoryGrade As Integer
    Dim targetColumnHeader As String
    Dim targetCol As Long
    Dim lastRowJeongsi As Long
    Dim lastColJeongSi As Long
    
    
    koreanHistoryGreade = 0
    koreanHistoryGrade = wsStudentData.Range(koreanHistoryGradeCellAddress).Value ' 한국사 등급 (E16 셀)
    
    ' 한국사 등급이 유효한 범위(1~9)인지 확인
    If koreanHistoryGrade < 1 Or koreanHistoryGrade > 9 Then
        MsgBox "대학 학과별 상세검색 시트의 한국사 등급(E16)이 1~9 범위를 벗어났습니다.", vbCritical
        Exit Sub
    End If
    
    
    ' 4. 한국사 등급에 맞는 컬럼 헤더 찾기 (예: "한1", "한2"...)
    targetColumnHeader = "한" & koreanHistoryGrade
    
    lastColJeongSi = wsJeongsi.Cells(1, Columns.Count).End(xlToLeft).Column ' 1행 헤더의 마지막 컬럼 찾기
    targetCol = 0 ' 초기화
    
    For i = 1 To lastColJeongSi
        If wsJeongsi.Cells(1, i).Value = targetColumnHeader Then ' 1행은 헤더라고 가정
            targetCol = i
            Exit For
        End If
    Next i
    
    If targetCol = 0 Then
        MsgBox "정시 시트에서 '" & targetColumnHeader & "' 컬럼을 찾을 수 없습니다.", vbExclamation
        wsStudentData.Range("F16").Value = "컬럼없음" ' 한국사 결과를 넣을 셀 F16으로 변경
        Exit Sub
    End If
    
    ' 5. 해당하는 한국사 점수 가져오기
    koreanHistoryScore = 0
    koreanHistoryScoreDisplay = 0

    koreanHistoryScore = wsJeongsi.Cells(jeongsiFoundRow, targetCol).Value
    koreanHistoryScoreDisplay = wsJeongsi.Cells(jeongsiFoundRow, targetCol).Value
    
    ' 6. '대학 학과별 상세검색' 시트에 결과 출력 (F16 셀에 출력하도록 변경)
'    wsStudentData.Range("F16").Value = koreanHistoryScore
    
    koreanHistoryScoreProcessType = ""
    
    If IsNumeric(koreanHistoryScore) Then
        koreanHistoryScoreProcessType = "변환점수"
        
    Else
        koreanHistoryScoreProcessType = ""
    End If
    
'    MsgBox "한국사 반영 점수를 성공적으로 가져왔습니다: " & koreanHistoryScore, vbInformation
    
End Sub

Sub ProcessInquiryScores()
    ' 변수 선언 (Dim)
    Dim inquirySubjectName(1 To 2) As String
    Dim inquiryStdScore(1 To 2) As Double
    Dim inquiryPercentile(1 To 2) As Double
    Dim inquiryGrade(1 To 2) As Integer
    Dim inquiryRows(1 To 2) As Integer
    Dim convertScoreResult(1 To 2) As Variant '국어, 수학 변호나 점수 결과 저장
    Dim convertScoreTypeResult(1 To 2) As String '대학 국어, 수학 반영 기준 예) 백분위, 표준점수, 등급, 변환표준점수
    
    Dim outputResult(1 To 2) As Variant ' 최종 결과 (점수 또는 텍스트)
    Dim outputCells(1 To 2) As Range   ' 결과를 출력할 셀 (F17, F18)
    Dim i As Integer ' 탐구 과목 루프 카운터
    Dim roundedPercentile As Integer
    
    
    Dim convertDBUnivCol As Long ' 대학별 탐구 변환표준점수표 DB 시트의 대학 컬럼
    Dim convertDBPercentileRow As Long
    Dim lastRowJeongsi As Long, lastColJeongSi As Long
    Dim lastRowConvertDB As Long, lastColConvertDB As Long
    
    Dim lookupUnivName As String ' 변환표준점수 DB에서 대학 컬럼을 찾을 때 사용할 이름 (고려대-과탐/사탐 등)
    Dim inquiryConvertScoreClassificationType As String
    
    ' For 루프 내에서 사용될 변수들은 미리 선언
    Dim col As Long
    Dim rowCSAT As Long, rowMock As Long
    Dim lastChar As String
'    Dim searchSatExamYearSubjectGradeString As String
'    Dim searchSatExamYearSubjectGradeFoundRange As Range
'    Dim searchSatExamYearSubjectGradeRange As Range
'    Dim searchSatExamYearSubjectGradeRow As String
    
    '해당 시험이 표준점수를 비교대상 정시의 표준점수로 변환하기 위한 변수정의
    Dim inputStdScore As Double
    Dim inputPercentile As Double
    Dim inputGrade As Integer
    Dim minMockScore As Double, maxMockScore As Double
    Dim minCSATScore As Double, maxCSATScore As Double
    Dim relativePos As Double
    Dim convertedScore As Double
    Dim foundCSAT As Boolean, foundMock As Boolean
    Dim lastRowCSAT As Long, lastRowMock As Long
    
    Dim alertMsg As String
    
    Dim searchSatString As String
    Dim foundSatRange As Range
    Dim searchSatRange As Range
    Dim foundSatRow As String
    Dim validateErrorFlag As Integer
    
    ' ... 기존 변수 선언에 추가 ...
    Dim inquiryMaxString As String
    Dim foundInquiryMaxRange As Range
    Dim foundInquiryMaxRow As Long
    
    ' 1. 시트 설정
    
    ' 2. '대학 학과별 상세검색' 시트에서 공통 정보 가져오기
    '이 프로시저에서 필요한 배열의 초기화
     For i = 1 To 2 ' 첫 번째 차원 (0 또는 1)
        inquirySubjectName(i) = Empty
        inquiryStdScore(i) = 0
        inquiryPercentile(i) = 0
        inquiryGrade(i) = 0
'        Set outputCells(i) = Nothing
        convertScoreResult(i) = 0
        convertScoreTypeResult(i) = Empty
        inquiryRows(i) = 0
     Next i
     
     
    lastChar = ""
    alertMsg = "검증 결과 - 표준점수 범위 벗어난 과목:" & vbNewLine
    validateErrorFlag = 0
     
    lastRowCSAT = wsDB_CSAT.Cells(wsDB_CSAT.Rows.Count, "A").End(xlUp).Row
    lastRowMock = wsDB_Mock.Cells(wsDB_Mock.Rows.Count, "A").End(xlUp).Row
     
    ' --- (초기화 끝) ---
    
    inquiryConvertScoreClassificationType = wsJeongsi.Cells(jeongsiFoundRow, inquiryConvertScoreClassificationCellAddress).Value
    
    
    Set searchSatRange = wsDB_CSAT.Columns(6)
    
    ' 6. 탐구 1, 2 과목별로 점수 처리
    For i = 1 To 2 ' 1은 탐구1 (17행), 2는 탐구2 (18행)
        ' '대학 학과별 상세검색' 시트에서 각 탐구 과목 정보 읽어오기
        inquirySubjectName(i) = Trim(wsStudentData.Cells(inquiryMileStoneRowNo + i, "B").Value) ' B17, B18 (과목명), Trim으로 공백 제거
        
        ' 탐구 과목이 입력되었는지 확인
        If inquirySubjectName(i) = "" Then
'            outputCells(i).Value = "과목없음" ' 탐구 과목이 없으면 '과목없음'으로 표시
            GoTo NextSubject ' 다음 탐구 과목으로 바로 넘어감
            
        '2022개정 교육과정의 학생이 통합사회, 통합과학으로 과목을 선택하였을 때, 이전 학년도의 수능 선택과목으로 강제 치환
        ElseIf inquirySubjectName(i) = "사회탐구" And _
                (examYearClassification = "2024학년도 정시" Or _
                examYearClassification = "2025학년도 정시" Or _
                 examexamYearClassification = "2026학년도 정시" Or _
                 examYearClassification = "2027학년도 정시" _
                 ) Then
        
            inquirySubjectName(i) = "사회·문화"    '2022개정교육과정 학생이
            
        ElseIf inquirySubjectName(i) = "과학탐구" And _
                (examYearClassification = "2024학년도 정시" Or _
                examYearClassification = "2025학년도 정시" Or _
                 examYearClassification = "2026학년도 정시" Or _
                 examYearClassification = "2027학년도 정시" _
                 ) Then
            inquirySubjectName(i) = "지구과학Ⅰ"
        
        End If

        inquiryStdScore(i) = wsStudentData.Cells(inquiryMileStoneRowNo + i, "C").Value
        inquiryPercentile(i) = wsStudentData.Cells(inquiryMileStoneRowNo + i, "D").Value
        inquiryGrade(i) = wsStudentData.Cells(inquiryMileStoneRowNo + i, "E").Value
        inquiryRows(i) = inquiryMileStoneRowNo + i
        
        
        '========================================
        ' 추가: 탐구 과목 최고점을 1등급 고정으로 조회
        '========================================
        inquiryMaxString = examYearClassification & inquirySubjectName(i) & 1
        Set foundInquiryMaxRange = searchSatRange.Find(What:=inquiryMaxString, _
                                                       LookIn:=xlValues, _
                                                       LookAt:=xlWhole, _
                                                       SearchOrder:=xlByRows, _
                                                       SearchDirection:=xlNext, _
                                                       MatchCase:=False)
        If Not foundInquiryMaxRange Is Nothing Then
            foundInquiryMaxRow = foundInquiryMaxRange.Row
            If i = 1 Then subjectSatStdScoreMax(3) = wsDB_CSAT.Cells(foundInquiryMaxRow, "E").Value  '탐구1 최고점
            If i = 2 Then subjectSatStdScoreMax(4) = wsDB_CSAT.Cells(foundInquiryMaxRow, "E").Value  '탐구2 최고점
        End If
        '========================================
        
        
        Select Case jeongsiScoreProcessType
            Case "백분위"
                convertScoreResult(i) = inquiryPercentile(i)
            Case "표준점수"
'                searchSatExamYearSubjectGradeString = examYearClassification & " " & inquirySubjectName(i) & " " & inquiryGrade(i)
'                Set searchSatExamYearSubjectGradeFoundRange = searchSatExamYearSubjectGradeRange.Find(What:=searchSatExamYearSubjectGradeString, _
'                                                                                                        LookIn:=xlValues, _
'                                                                                                        LookAt:=xlWhole, _
'                                                                                                        SearchOrder:=xlByRows, _
'                                                                                                        SearchDirection:=xlNext, _
'                                                                                                        MatchCase:=False)
'
'                If Not searchSatExamYearSubjectGradeFoundRange Is Nothing Then
'
'                    ' 찾은 셀이 있다면 해당 행 번호를 변수에 저장
'                    searchSatExamYearSubjectGradeRow = searchSatExamYearSubjectGradeFoundRange.Row
'                Else
'                    MsgBox "'" & targetChar & "' 데이터를 찾을 수 없습니다."
'                    Exit Sub
'                End If
'                For rowMock = 2 To lastRowMock
'                    '=======================================================
'                    '탐구과목
'                    '모의고사 해당 과목 등급의 최대, 최솟값 찾기
'                    '=======================================================
'                    If wsDB_Mock.Cells(rowMock, "A").Value = satYearVal And _
'                        wsDB_Mock.Cells(rowMock, "B").Value = satGradeVal And _
'                        wsDB_Mock.Cells(rowMock, "C").Value = satMonthVal And _
'                        wsDB_Mock.Cells(rowMock, "D").Value = inquirySubjectName(i) And _
'                        wsDB_Mock.Cells(rowMock, "E").Value = inquiryGrade(i) Then
'                        minMockScore = wsDB_Mock.Cells(rowMock, "F").Value
'                        maxMockScore = wsDB_Mock.Cells(rowMock, "G").Value
'                        foundMock = True
'                        Exit For
'                    End If
'                Next rowMock
                
                If rowDB_MockInquiry(i) > 0 Then

                        minMockScore = wsDB_Mock.Cells(rowDB_MockInquiry(i), "F").Value
                        maxMockScore = wsDB_Mock.Cells(rowDB_MockInquiry(i), "G").Value
                        foundMock = True
                End If
                
                
                
                searchSatString = examYearClassification & inquirySubjectName(i) & inquiryGrade(i)
                Set foundSatRange = searchSatRange.Find(What:=searchSatString, _
                                                    LookIn:=xlValues, _
                                                    LookAt:=xlWhole, _
                                                    SearchOrder:=xlByRows, _
                                                    SearchDirection:=xlNext, _
                                                    MatchCase:=False)
        
                If Not foundSatRange Is Nothing Then
        
                    ' 찾은 셀이 있다면 해당 행 번호를 변수에 저장
                    foundSatRow = foundSatRange.Row
                    minCSATScore = wsDB_CSAT.Cells(foundSatRow, "D").Value
                    maxCSATScore = wsDB_CSAT.Cells(foundSatRow, "E").Value
                    foundCSAT = True
        
                Else
                    MsgBox i & "번 째 탐구과목의 정보를 찾을 수 없습니다." & vbCrLf, vbInformation
                    Exit Sub
                End If
                
'                If foundCSAT Then
'                    If inputStdScore < minCSATScore Or inputStdScore > maxCSATScore Then
'                        ' 입력한 표준점수가 범위를 벗어났을 경우
'                        alertMsg = alertMsg & subjectsName(i) & " - 등급 " & inputGrade & ", 표준점수 " & inputStdScore & " (기준 " & minCSATScore & " ~ " & maxCSATScore & ")" & vbNewLine
'                        validateErrorFlag = validateErrorFlag + 1
'                    End If
'                End If
                
                
                
                
'                For rowCSAT = 2 To lastRowCSAT
'                    '=======================================================
'                    '탐구과목
'                    '수능 해당 과목 등급의 최대, 최솟값 찾기
'                    '=======================================================
'                    If wsDB_CSAT.Cells(rowCSAT, "A").Value = examYearClassification And _
'                        wsDB_CSAT.Cells(rowCSAT, "B").Value = inquirySubjectName(i) And _
'                        wsDB_CSAT.Cells(rowCSAT, "C").Value = inquiryGrade(i) Then
'                        minCSATScore = wsDB_CSAT.Cells(rowCSAT, "D").Value
'                        maxCSATScore = wsDB_CSAT.Cells(rowCSAT, "E").Value
'                        foundCSAT = True
'                        Exit For
'                    End If
'                Next rowCSAT
                
                
                
                
                If foundMock And foundCSAT Then
                    If maxMockScore = minMockScore Then
                        relativePos = 0
                    Else
                        relativePos = (inquiryStdScore(i) - minMockScore) / (maxMockScore - minMockScore)
                    End If
                    
                    If relativePos < 0 Then relativePos = 0
                    If relativePos > 1 Then relativePos = 1
                    
                    convertedScore = minCSATScore + relativePos * (maxCSATScore - minCSATScore)
                    convertScoreResult(i) = CLng(Application.WorksheetFunction.Round(convertedScore, 0))
                    convertScoreTypeResult(i) = "표준점수"
                Else
                    If worksheetName = "대학 학과별 상세검색" Then
                        wsStudentData.Cells(inquiryRows(i), "F").Value = "기준없음"
                    End If
                End If
                
                '2024학년도, 2025학년도 약대 의대는 탐구는 백분위로....
                If targetUniversity = "단국대학교_천안" Then
                    convertScoreResult(i) = inquiryPercentile(i)
'                Else
'
'                    convertScoreResult(i) = inquiryStdScore(i)
                End If
            Case "등급"
                convertScoreResult(i) = inquiryGrade(i)
            Case "변환표준점수"
                ' 고려대, 시립대, 성균관대학교 특수 처리 로직!
                lookupUnivName = targetUniversity ' 기본적으로 원래 대학명 사용
                
                '=============================================
                '2024학년도 정시
                '============================================
                If examYearClassification = "2024학년도 정시" Then
                
                    '==================================================
                    'ABC 구분자가 있는 경우와 없는 경우
                    '==================================================
                    If Len(inquiryConvertScoreClassificationType) > 0 Then
                    
                        '===========================================================
                        '광운대학교, 동국대학교, 성균관대학교인 경우와 그 외 경우
                        '===========================================================

                        If targetUniversity = "광운대학교" Or targetUniversity = "한양대학교" Then
                            
                            If inquiryConvertScoreClassificationType = "A" Then
                            
                                lookupUnivName = lookupUnivName & "-" & inquiryConvertScoreClassificationType
                            ElseIf inquiryConvertScoreClassificationType = "B" Then
                            
                                '광운대학교는 B 타입이면서 과탐일때와  B 타입이면서 과탐이 아니면 사탐으로 돌리자.
                                
                                If targetUniversity = "광운대학교" Then
                                    If Len(inquirySubjectName(i)) > 0 Then
                                        lastChar = Right(inquirySubjectName(i), 1)
                                        If lastChar = "Ⅰ" Or lastChar = "Ⅱ" Then ' 과탐인 경우
                                            lookupUnivName = targetUniversity & "-" & inquiryConvertScoreClassificationType
                                        Else
                                            lookupUnivName = targetUniversity & "-" & "A"   'B type이나 사탐인 경우  A로 바꾼다
                                        End If
                                    End If
                                Else
                                    lookupUnivName = lookupUnivName & "-" & inquiryConvertScoreClassificationType
                                End If
                            End If


                        ElseIf targetUniversity = "동국대학교" Then
                        
                            If inquiryConvertScoreClassificationType = "C" Then
                            
                                If Len(inquirySubjectName(i)) > 0 Then
                                    lastChar = Right(inquirySubjectName(i), 1)
                                    If lastChar = "Ⅰ" Or lastChar = "Ⅱ" Then ' 과탐인 경우
                                        lookupUnivName = targetUniversity & "-" & inquiryConvertScoreClassificationType & "-과탐"
                                    Else ' 사탐인 경우 (로마자 1,2가 아니면 사탐으로 간주)
                                        lookupUnivName = targetUniversity & "-" & inquiryConvertScoreClassificationType & "-사탐"
                                    End If
                                End If
                            Else
                            
                                lookupUnivName = targetUniversity & "-" & inquiryConvertScoreClassificationType
                            
                            End If
                                                        
                        ElseIf targetUniversity = "성균관대학교" Then
                                
                            If Len(inquirySubjectName(i)) > 0 Then
                                lastChar = Right(inquirySubjectName(i), 1)
                                If lastChar = "Ⅰ" Or lastChar = "Ⅱ" Then ' 과탐인 경우
                                    lookupUnivName = targetUniversity & "-" & inquiryConvertScoreClassificationType & "-과탐"
                                Else ' 사탐인 경우 (로마자 1,2가 아니면 사탐으로 간주)
                                    lookupUnivName = targetUniversity & "-" & inquiryConvertScoreClassificationType & "-사탐"
                                End If
                            End If
                            
                        '==========================================================================
                        '숭실대학교, 인하대학교
                        '==========================================================================
                        Else
                    
                            lookupUnivName = targetUniversity & "-" & inquiryConvertScoreClassificationType
                        
                        End If
                    
                    '=========================================================================
                    'ABC 구분자가 없는 경우
                    '=========================================================================
                    Else
                        If targetUniversity = "경북대학교" Or _
                            targetUniversity = "고려대학교" Or _
                            targetUniversity = "서울시립대학교" Or _
                            targetUniversity = "세종대학교" Or _
                            targetUniversity = "아주대학교" Or _
                            targetUniversity = "연세대학교_미래" Or _
                            targetUniversity = "전북대학교" Or _
                            targetUniversity = "고려대학교_세종" Or _
                            targetUniversity = "이화여자대학교" Then
                        
                            If Len(inquirySubjectName(i)) > 0 Then
                                lastChar = Right(inquirySubjectName(i), 1)
                                If lastChar = "Ⅰ" Or lastChar = "Ⅱ" Then ' 과탐인 경우
                                    lookupUnivName = targetUniversity & "-과탐"
                                Else ' 사탐인 경우 (로마자 1,2가 아니면 사탐으로 간주)
                                    lookupUnivName = targetUniversity & "-사탐"
                                End If
                            End If
                        
                        End If
                    End If
                
                
                '=======================================================
                '2025학년도 정시
                '=======================================================
                ElseIf examYearClassification = "2025학년도 정시" Then
                    If targetUniversity = "고려대학교" Or _
                        targetUniversity = "고려대학교_세종" Or _
                        targetUniversity = "서울시립대학교" Or _
                        targetUniversity = "성균관대학교" Or _
                        targetUniversity = "아주대학교" Or _
                        targetUniversity = "전북대학교" Or _
                        targetUniversity = "경북대학교" Then
                        ' 과목명 마지막 글자가 로마자 1,2 (I, II) 인지 확인
                        If Len(inquirySubjectName(i)) > 0 Then
                            lastChar = Right(inquirySubjectName(i), 1)
                            If lastChar = "Ⅰ" Or lastChar = "Ⅱ" Then ' 과탐인 경우
                                lookupUnivName = targetUniversity & "-과탐"
                            Else ' 사탐인 경우 (로마자 1,2가 아니면 사탐으로 간주)
                                lookupUnivName = targetUniversity & "-사탐"
                            End If
                        End If
                    
                     ' 가톨릭대학교 특수 처리 로직 (학과 구분)
                    ElseIf targetUniversity = "가톨릭대학교" Then
                        If targetMajor = "의예과" Or targetMajor = "약학과" Then
                            lookupUnivName = targetUniversity & "-약의"
                        ElseIf targetMajor = "간호학과" Then
                            lookupUnivName = targetUniversity & "-간호"
                        End If
                        ' 참고: 가톨릭대학교는 과탐/사탐 구분 없이 학과로만 추가 접미사가 붙는다고 가정
                    End If
                
                
                ElseIf examYearClassification = "2026학년 정시" Then
                
                
                
                End If
                
                
                
                ' 5. '대학별 탐구 변환표준점수표 DB' 시트에서 (변경될 수도 있는) 대학의 컬럼 찾기
                lastColConvertDB = wsConvertDB.Cells(1, Columns.Count).End(xlToLeft).Column
                convertDBUnivCol = 0
                
                ' C열부터 대학명 헤더 (A:수능구분, B:백분위)
                For col = 3 To lastColConvertDB
                    If wsConvertDB.Cells(1, col).Value = lookupUnivName Then
                        convertDBUnivCol = col
                        Exit For
                    End If
                Next col
                
                If convertDBUnivCol = 0 Then
                    convertScoreResult(i) = lookupUnivName & " DB컬럼없음"
                    GoTo NextSubject ' 이 과목은 건너뛰고 다음 과목으로
                End If
                
                roundedPercentile = CInt(Application.WorksheetFunction.Round(inquiryPercentile(i), 0))
                
                ' 백분위 값 유효성 검사 추가 (0이하 값은 백분위 정보 없다고 판단)
                If inquiryPercentile(i) <= 0 Then
                    convertScoreTypeResult(i) = "백분위 정보없음"
                    GoTo NextSubject
                End If

                lastRowConvertDB = wsConvertDB.Cells(Rows.Count, "B").End(xlUp).Row
                convertDBPercentileRow = 0
                
                For rowCSAT = 2 To lastRowConvertDB
                    If wsConvertDB.Cells(rowCSAT, "B").Value = examYearClassification And _
                       wsConvertDB.Cells(rowCSAT, "C").Value = roundedPercentile Then
                        convertDBPercentileRow = rowCSAT
                       Exit For
                    End If
                Next rowCSAT
                
                If convertDBPercentileRow <> 0 Then
                    convertScoreResult(i) = wsConvertDB.Cells(convertDBPercentileRow, convertDBUnivCol).Value
                    
                    '아주대학교, 인하대학교는 대학환산점수를 낼 때,
                    ' 2024학년도 2025학년도에 변환표준점수의 최댓값을 이용한다.
                    '여기에서 미리 저장을 해두자.
                    'i=1이면 subjectSatStdScoreMax(3), i=2이면 subjectSatStdScoreMax(4)
                    
                    If targetUniversity = "아주대학교" Or _
                        targetUniversity = "인하대학교" Then
                    
                        subjectSatStdScoreMax(i + 2) = wsConvertDB.Cells(2, convertDBUnivCol).Value
                    End If
                Else
                    ' 변환표 DB에서 해당 변환점수를 찾지 못한 경우 입력된 백분위를 대신 사용
                    convertScoreResult(i) = inquiryPercentile(i)
                End If
            Case Else
                convertScoreTypeResult(i) = "처리유형오류" ' AC열 값이 예상 외인 경우
        End Select
        
        ' 결과 출력
'        outputCells(i).Value = outputResult(i)
        If examYearClassification = "2025학년도 정시" And targetUniversity = "단국대학교_천안" Then
            jeongsiScoreProcessType = "백분위"
        End If
        convertScoreTypeResult(i) = jeongsiScoreProcessType
        
NextSubject: ' GoTo 문을 위한 레이블
    Next i
    
    inquiry1SubjectName = ""
    inquiry1Score = 0
    inquiry1ScoreDisplay = 0
    inquiry1ScoreProcessType = ""
    inquiry2SubjectName = ""
    inquiry2Score = 0
    inquiry2ScoreDisplay = 0
    inquiry2ScoreProcessType = ""
    
    For i = 1 To 2
    
        If i = 1 Then
            inquiry1SubjectName = inquirySubjectName(1)
            inquiry1Score = convertScoreResult(i)
            inquiry1ScoreDisplay = convertScoreResult(i)
            inquiry1ScoreProcessType = convertScoreTypeResult(i)
        ElseIf i = 2 Then
            inquiry2SubjectName = inquirySubjectName(2)
            inquiry2Score = convertScoreResult(i)
            inquiry2ScoreDisplay = convertScoreResult(i)
            inquiry2ScoreProcessType = convertScoreTypeResult(i)
        End If
        
    Next i

    


End Sub

Sub SecondLanguage()
    Dim lastColJeongSi As Long
    Dim targetCol As Long
    Dim targetColumnHeader As String
    
    
    If targetUniversity = "서울대학교" Then
    
        If IsNumeric(wsStudentData.Range(secondLanguageGradeCellAddress).Value) Then
        
            lastColJeongSi = wsJeongsi.Cells(1, Columns.Count).End(xlToLeft).Column ' 1행 헤더의 마지막 컬럼 찾기
            targetCol = 0 ' 초기화
            targetColumnHeader = "제2한" & wsStudentData.Range(secondLanguageGradeCellAddress).Value
            
            For c = 1 To lastColJeongSi
                If wsJeongsi.Cells(1, c).Value = targetColumnHeader Then ' 1행은 헤더라고 가정
                    targetCol = c
                    Exit For
                End If
            Next c
            
            
            If targetCol = 0 Then
            
            
                '대학별 검색에서 제2외국어 관련 메세지를 한 번만 뿌리 하는 로직
                If secondLanguageFlag = 0 Then
                    MsgBox "서울대학교 정시는 제2외국어/한문 성적이 필수입니다." & vbCrLf & "제2외국어/한문 성적의 입력 값을 확인해 주세요.", vbExclamation
                    secondLanguageFlag = secondLanguageFlag + 1
                End If
                Exit Sub
            End If
            
            secondLanguageScore = 0
            secondLanguageScoreDisplay = 0
            secondLanguageScore = wsJeongsi.Cells(jeongsiFoundRow, targetCol).Value
            secondLanguageScoreDisplay = wsJeongsi.Cells(jeongsiFoundRow, targetCol).Value


            secondLanguageScoreProcessType = ""
            
            If IsNumeric(secondLanguageScore) And secondLanguageScore > 0 Then
            
               secondLanguageScoreProcessType = "변환점수"
               
            Else
                secondLanguageScoreProcessType = ""
            End If
        End If
    End If
    

End Sub

Sub ResultCellClear()

    Dim ws As Worksheet, wsDB As Worksheet
    Dim wsNmae As String
    ' "대학 학과별 상세검색" 시트와 "모의고사학생입력성적유효범위" 시트 설정
    
    wsName = ThisWorkbook.ActiveSheet.Name
    If wsName = "대학 학과별 상세검색" Then
        Set ws = ThisWorkbook.Sheets("대학 학과별 상세검색")
        ws.Range("F13:Ｈ19").ClearContents
        ws.Range("D8:Ｈ8").ClearContents

    ElseIf wsName = "대학별검색" Then
    
        Set ws = ThisWorkbook.Sheets("대학별검색")
        
        ws.Range("A17:L10000").Borders.LineStyle = xlNone

        ws.Range("A17:L10000").ClearContents
    
    End If
End Sub




'가산점에 관한 정보를 저장

Sub BonusProcess()
    Dim bonusData As Double
    Dim combineResult As String '서울대 과탐 Ⅰ,Ⅱ 저장 변수
    Dim inquiry1LastChar As String
    Dim inquiry2LastChar As String
    
    Dim i As Long
    
    
    
    mathSubjectName = ""
    mathSubjectName = wsStudentData.Range(mathSubjectNameCellAddress)
    
    If mathSubjectName = "공통수학" And _
                (examYearClassification = "2024학년도 정시" Or _
                 examYearClassification = "2025학년도 정시" Or _
                 examYearClassification = "2026학년도 정시" Or _
                 examYearClassification = "2027학년도 정시" _
                 ) Then
        
            mathSubjectName = "미적분"    '2022개정교육과정 학생이
    End If
    
    ' 가산점에 대한 절차 수행
    '보너스 첫번째 행이 Y이면 가산점이 있다. bonusStartColNo=80으로 CB열
    If wsJeongsi.Cells(jeongsiFoundRow, bonusStartColNo) = "Y" Then
        
        inquiry1LastChar = ""
        inquiry2LastChar = ""
        
        
        inquiry1LastChar = Right(inquiry1SubjectName, 1)
        inquiry2LastChar = Right(inquiry2SubjectName, 1)
        
        ' 위 if 절에서 i=0를 수행했기에  1부터 시작함.
        For i = 1 To bonusColCount - 1
            
            '가산점 값을 아래의 if 절에서 할당, 활용할 변수값 초기화
            bonusCase1 = 0
            bonusCase2 = 0
            bonusCase3 = 0
            bonusCase4 = 0
            bonusCase7 = 0
            bonusCase8 = 0
            bonusCase9 = 0
            bonusCase10 = 0
            
            combineResult = ""
            bonusData = 0
            
            '해당 칸에 가산점관련 정보가 있는지 판단
            If IsNumeric(wsJeongsi.Cells(jeongsiFoundRow, bonusStartColNo + i).Value) And _
                wsJeongsi.Cells(jeongsiFoundRow, bonusStartColNo + i).Value > 0 Then
                
                    bonusData = wsJeongsi.Cells(jeongsiFoundRow, bonusStartColNo + i).Value
                    
                    Select Case i
                        
                        Case 1  '확통가산
                                'CM열
                                '확률과 통계 가산
                                ' 첫 번째 가산점 정보에 대한 처리 절차
                                ' 특정한 점수가 가산되는 경우이다
                            If mathSubjectName = "확률과 통계" Then
                            
                                    mathScore = mathScore * (1 + bonusData)
                            End If
                            
                            Debug.Print targetUniversity & " Case 2: " & mathSubjectName & "Processing bonus data " & bonusData
                            Debug.Print targetUniversity & " Case 1: Processing bonus data " & bonusData
                                                        
                        Case 2  '미적분가산
                                'CN열
                                '두 번째 가산점 정보에 대한 처리 절차
                            If mathSubjectName = "미적분" Then
                                If targetUniversity = "동덕여자대학교" And targetMajor = "자연정보융합학부" Then
                                
                                    If mathScore > koreanScore Then
                                        
                                        mathScore = mathScore * (1 + bonusData)
                                    End If
                                ElseIf targetUniversity = "경상국립대학교" Then
                                
                                    bonusCase2 = bonusData
                                Else
                                
                                    mathScore = mathScore * (1 + bonusData)
                                End If
                            End If
                            
                            Debug.Print targetUniversity & " Case 2: " & mathSubjectName & "Processing bonus data " & bonusData
                            
                            
                        Case 3  '기하가산
                                'CO열
                                ' 세 번째 가산점 정보에 대한 처리 절차
                            If mathSubjectName = "기하" Then
                            
                                If targetUniversity = "동덕여자대학교" And targetMajor = "자연정보융합학부" Then
                                
                                    If mathScore > koreanScore Then
                                        
                                        mathScore = mathScore * (1 + bonusData)
                                        
                                    End If
                                ElseIf targetUniversity = "경상국립대학교" Then
                                
                                    bonusCase3 = bonusData
                                Else
                                
                                    mathScore = mathScore * (1 + bonusData)
                                End If
                            End If
                                                        
                            Debug.Print targetUniversity & " Case 3: " & mathSubjectName & "Processing bonus data " & bonusData
                        
                        
                        Case 4  '과탐가산
                                ' 네 번째 가산점 정보에 대한 처리 절차
                                'CP열
                                '과탐가산
                                '2025학년도 성신여대는 두 과목 다 과학일 때는 큰 과목에게 가산점을 준다.
                            
                            bonusFlag = 0
                            If targetUniversity = "성신여자대학교" Then
                                        
                                    If inquiry1LastChar = "Ⅰ" Or inquiry1LastChar = "II" Then
                                            Debug.Print "인증된 과목명입니다: " & inquiry1SubjectName
                                            bonusFlag = bonusFlag + 1
                                       
                                    End If
            

                                    If inquiry2LastChar = "Ⅰ" Or inquiry2LastChar = "II" Then
                                        ' 해당 과목일 경우 수행할 작업
                                        
                                            ' 해당 과목일 경우 수행할 작업
                                            Debug.Print "인증된 과목명입니다: " & inquiry2SubjectName
                                            bonusFlag = bonusFlag + 3
                                        
                                    End If
                                    
                                    If bonusFlag = 1 Then
                                    
                                            bonusCase4 = inquiry1Score * bonusData
                                            
                                    ElseIf bonusFlag = 3 Then
                                    
                                            bonusCase4 = inquiry2Score * bonusData
                                    
                                    ElseIf bonusFlag = 4 Then
                                    
                                            If inquiry1Score > inquiry2Score Then
                                            
                                                bonusCase4 = inquiry1Score * bonusData
                                            Else
                                                bonusCase4 = inquiry2Score * bonusData
                                            End If
                                    
                                    End If
                                    
                            ElseIf targetUniversity = "건국대학교_글로컬" Then
                                        
                                    If inquiry1LastChar = "Ⅰ" Or inquiry1LastChar = "II" Then
                                            bonusFlag = bonusFlag + 1
                                       
                                    End If
            

                                    If inquiry2LastChar = "Ⅰ" Or inquiry2LastChar = "II" Then
                                        ' 해당 과목일 경우 수행할 작업
                                        
                                            ' 해당 과목일 경우 수행할 작업
                                            bonusFlag = bonusFlag + 3
                                        
                                    End If
                                    
                                    If bonusFlag = 1 Then
                                    
                                            inquiry1Score = inquiry1Score * (1 + bonusData)
                                            
                                    ElseIf bonusFlag = 3 Then
                                    
                                            inquiry2Score = inquiry2Score * (1 + bonusData)
                                    
                                    ElseIf bonusFlag = 4 Then
                                    
                                            If inquiry1Score > inquiry2Score Then
                                            
                                                inquiry1Score = inquiry1Score * (1 + bonusData)
                                            Else
                                                inquiry2Score = inquiry2Score * (1 + bonusData)
                                            End If
                                    
                                    End If
                            ElseIf targetUniversity = "을지대학교" Then
                            
                                If inquiry1LastChar = "Ⅰ" Or inquiry1LastChar = "II" Then
                                    bonusFlag = bonusFlag + 1
                                    Debug.Print targetUniversity & " Case 4 " & " inquiry1SubjectName 과목명입니다: " & inquiry1SubjectName
                                    
                                End If
                                
            
                                If inquiry2LastChar = "Ⅰ" Or inquiry2LastChar = "II" Then
                                                                   
                                    bonusFlag = bonusFlag + 2
                                    Debug.Print targetUniversity & "Case 4 " & "  inquiry2SubjectName 과목명입니다: " & inquiry2SubjectName
                                    
                                End If
                                
                                Select Case bonusFlag
                                
                                    Case 1
                                    
                                        inquiry1Score = inquiry1Score * (1 + bonusData)
                                        
                                    Case 2
                                        inquiry2Score = inquiry2Score * (1 + bonusData)
                                        
                                End Select
                                
                            ElseIf targetUniversity = "경희대학교" Then
                            
                                If inquiry1LastChar = "Ⅰ" Or inquiry1LastChar = "II" Then
                                
                                    
                                        inquiry1Score = inquiry1Score + bonusData
                                    Debug.Print targetUniversity & " inquiry1SubjectName 과목명입니다: " & inquiry1SubjectName
                            
                                End If
                                
                                If inquiry2LastChar = "Ⅰ" Or inquiry2LastChar = "II" Then
                                
                                    
                                        inquiry2Score = inquiry2Score + bonusData
                                    Debug.Print targetUniversity & " inquiry2SubjectName 과목명입니다: " & inquiry2SubjectName
                            
                                End If
                                
                                Debug.Print "Case 4: Processing bonus data " & bonusData
                            
                            Else
                            
                                If inquiry1LastChar = "Ⅰ" Or inquiry1LastChar = "II" Then
                                
                                    
                                    inquiry1Score = inquiry1Score * (1 + bonusData)
                                        
                                    Debug.Print targetUniversity & " inquiry1SubjectName 과목명입니다: " & inquiry1SubjectName
                            
                                End If
                                
                                If inquiry2LastChar = "Ⅰ" Or inquiry2LastChar = "II" Then
                                
                                    
                                    inquiry2Score = inquiry2Score * (1 + bonusData)
                                     
                                    Debug.Print targetUniversity & " inquiry2SubjectName 과목명입니다: " & inquiry2SubjectName
                            
                                End If
                                
                                Debug.Print "Case 4: Processing bonus data " & bonusData
                            
                            
                            End If
                            
                        Case 5  '사탐가산
                                'CQ열
                                '사탐가산
                                ' 다섯 번째 가산점 정보에 대한 처리 절차
                                '2025학년도 서울시립대는 모두 사탐일 때 가산점
                            If targetUniversity = "서울시립대학교" Then
                            
                                If inquiry1LastChar <> "Ⅰ" And inquiry1LastChar <> "II" And _
                                    inquiry2LastChar <> "Ⅰ" And inquiry2LastChar <> "II" Then
                                        
                                        inquiry1Score = inquiry1Score * (1 + bonusData)
                                        
                                        inquiry2Score = inquiry2Score * (1 + bonusData)
                                        
                                        ' 해당 과목일 경우 수행할 작업
                                        Debug.Print "인증된 과목명입니다: " & inquiry1SubjectName
                                End If
                            ElseIf targetUniversity = "경희대학교" Then
                            
                                If inquiry1LastChar <> "Ⅰ" And inquiry1LastChar <> "II" Then
                                        
                                            inquiry1Score = inquiry1Score + bonusData
                                        
                                        Debug.Print targetUniversity & " Case 5 " & "inquiry1SubjectName 과목명입니다: " & inquiry1SubjectName & "가산점" & bonusData
                                
                                End If
                                
                                If inquiry2LastChar <> "Ⅰ" And inquiry2LastChar <> "II" Then
                                
                                        
                                            inquiry2Score = inquiry2Score + bonusData
                                        
                                        Debug.Print targetUniversity & " Case 5 " & "inquiry1SubjectName 과목명입니다: " & inquiry1SubjectName & "가산점" & bonusData
                                End If
                            
                            Else
                            
                                If inquiry1LastChar <> "Ⅰ" And inquiry1LastChar <> "II" Then
                                        
                                    inquiry1Score = inquiry1Score * (1 + bonusData)
                                    Debug.Print targetUniversity & " Case 5 " & "inquiry1SubjectName 과목명입니다: " & inquiry1SubjectName & "가산점" & bonusData
                                
                                End If
                                
                                If inquiry2LastChar <> "Ⅰ" And inquiry2LastChar <> "II" Then
                                
                                        
                                    inquiry2Score = inquiry2Score * (1 + bonusData)
                                    Debug.Print targetUniversity & " Case 5 " & "inquiry1SubjectName 과목명입니다: " & inquiry1SubjectName & "가산점" & bonusData
                            
                                End If
                            End If
                            
                        Case 6  '과탐2개
                                'CR열
                                ' 여섯 번째 가산점 정보에 대한 처리 절차
                                '2025학년도 서울시립대학교는 모두 과탐일 때 가산점
                                
                            If (inquiry1LastChar = "Ⅰ" Or inquiry1LastChar = "Ⅱ") And _
                                (inquiry2LastChar = "Ⅰ" Or inquiry2LastChar = "Ⅱ") Then
                                
                                    inquiry1Score = inquiry1Score * (1 + bonusData)
                                    inquiry2Score = inquiry2Score * (1 + bonusData)
                                    
                                    Debug.Print "인증된 과목명입니다: " & inquiry1SubjectName
                                    Debug.Print "인증된 과목명입니다: " & inquiry2SubjectName
                                    
                            End If
                            
                        
                        Case 7  '과탐Ⅰ,Ⅰ+Ⅰ
                                'CS열
                                '일곱 번째 가산점 정보에 대한 처리 절차
                                '2025학년도 을지대학교에서만 일어나는 일로 확인.
                            If targetUniversity = "경상국립대학교" Then
                            
                                combineResult = inquiry1LastChar & inquiry2LastChar
                            
                                If combineResult = "ⅠⅠ" Then
                                    bonusCase7 = bonusData
                                End If
                            
                            End If
                            
                        
                         Case 8 '과Ⅱ
                                'CT열
                                ' 여덟 번째 가산점 정보에 대한 처리 절차
                                '서강대학교, 단국대학교_천안의 의예과, 치의예과, 약학과
                            If targetUniversity = "서강대학교" And _
                                (inquiry1LastChar = "Ⅱ" Or inquiry2LastChar = "Ⅱ") Then
                                                      
                                    bonusCase8 = bonusData
                                    Debug.Print targetUniversity & " Case 8: Processing bonus data " & bonusData
                            Else
                                   
                                If inquiry1LastChar = "Ⅱ" Then
                                    inquiry1Score = inquiry1Score * (1 + bonusData)
                                End If
                            
                                If inquiry2LastChar = "Ⅱ" Then
                                    inquiry2Score = inquiry2Score * (1 + bonusData)
                                End If
                            End If
                        
                        Case 9  '과탐 ⅠⅡ
                                'CU열
                                ' 아홉 번째 가산점 정보에 대한 처리 절차
                            If targetUniversity = "서울대학교" Then
                            
                                combineResult = inquiry1LastChar & inquiry2LastChar
                                
                                
                                If (combineResult = "ⅠⅡ") Or (combineResult = "ⅡⅠ") Then
                                
                                    bonusCase9 = 3
                                
                                End If
                            End If '서울대 처리 끝
                                
                            If targetUniversity = "한국교원대학교" Then
                                If (InStr(targetMajor, "물리") And InStr(inquiry1SubjectName, "물리학") And InStr(inquiry2SubjectName, "물리학")) Or _
                                    (InStr(targetMajor, "화학") And InStr(inquiry1SubjectName, "화학") And InStr(inquiry2SubjectName, "화학")) Or _
                                    (InStr(targetMajor, "생물") And InStr(inquiry1SubjectName, "생명과학") And InStr(inquiry2SubjectName, "생명과학")) Or _
                                    (InStr(targetMajor, "지구과학") And InStr(inquiry1SubjectName, "지구과학") And InStr(inquiry2SubjectName, "지구과학")) Then
                                    
                                    combineResult = inquiry1LastChar & inquiry2LastChar
                                    
                                    
                                    If (combineResult = "ⅠⅡ") Or (combineResult = "ⅡⅠ") Then
                                    
                                            If inquiry1Score > inquiry2Score Then
                                            
                                                inquiry1Score = inquiry1Score * (1 + bonusData)
                                            Else
                                                inquiry2Score = inquiry2Score * (1 + bonusData)
                                            End If
                                        
                                        Debug.Print targetUniversity & " Case 9 inquiry1SubjectName 과목명입니다: " & inquiry1SubjectName
                                        Debug.Print targetUniversity & " Case 9 inquiry2SubjectName 과목명입니다: " & inquiry2SubjectName
                                    End If
                                            
                                End If
                            End If '한국교원대 처리 끝
                            
                            If targetUniversity = "경상국립대학교" Then
                            
                                combineResult = inquiry1LastChar & inquiry2LastChar
                                
                                
                                If combineResult = "ⅠⅡ" Or _
                                    combineResult = "ⅡⅠ" Then
        
                                
                                        bonusCase9 = bonusData
                                
                                End If
                            
                            End If '경상국립대학교 처리 끝
                            
                            
                        Case 10 '과탐ⅡⅡ
                                'CV열
                            combineResult = inquiry1LastChar & inquiry2LastChar
                             
                             
                            If combineResult = "ⅡⅡ" Then
                             
                                 If targetUniversity = "서울대학교" Then
                                 
                                     bonusCase10 = 5
                                 ElseIf targetUniversity = "경상국립대학교" Then
                                     bonusCase10 = bonusData
                                 End If '서울대 처리 끝
                            End If
                        
                        Case 11 '물리
                                'CW열
                                '열한 번째 가산점 정보 처리 절차
                                '숙명여자대학교 신소재물리전공
                                If inquiry1SubjectName = "물리학Ⅰ" Or _
                                    inquiry1SubjectName = "물리학Ⅱ" Then
                                    
                                        inquiry1Score = inquiry1Score * (1 + bonusData)
                                        Debug.Print targetUniversity & " Case 11 inquiry1SubjectName 과목명입니다: " & inquiry1SubjectName
                                End If
        
                                If inquiry2SubjectName = "물리학Ⅰ" Or _
                                    inquiry2SubjectName = "물리학Ⅱ" Then
                            
                                       inquiry2Score = inquiry2Score * (1 + bonusData)
                                        Debug.Print targetUniversity & " Case 11 inquiry2SubjectName 과목명입니다: " & inquiry2SubjectName
                                End If
                                                            
                        Case 12
                                'CX열
                        
                        
                        Case 13
                                'CY열
                        
                    End Select
             End If
        Next i
    End If
    
End Sub





'반영비율과 점수를 배열로 같이 저장

Sub WeightRatioProcess() ' <-- 프로시저 이름

    ' 모듈 수준 변수들은 이 Sub 내부에서 다시 선언하거나 값을 재할당하지 않습니다.

'    Dim weightRatioStartColNo As Long     ' 검색을 시작할 열 번호 (AD열은 30번)
'    Dim weightRatioEndColNo As Long       ' 검색을 종료할 열 번호 (20개 열을 검사)
    Dim iCol As Long            ' 열을 순회하기 위한 루프 변수
    Dim position As Long        '성균관대학교 비율에 "(" ")"가 들어 있는지 확인
    

    Dim arrayIndex As Long      ' weightRatio 배열에 값을 채울 때 사용할 인덱스

    Dim cellValue As Double
    
    Dim arrMaxMidMin() As Double
    

    
    
    ' --- [weightRatio 배열 명시적 초기화 부분] ---
    ' 모듈 수준 변수 weightRatio는 선언 시 Empty로 자동 초기화되지만,
    ' 명시적인 초기화 과정을 Sub 내에서 한 번 더 수행하여 명확성을 높이고
    ' 혹시 모를 이전 값의 잔여를 완전히 비웁니다.
    Dim r As Long, c As Long
    For r = LBound(weightRatio, 1) To UBound(weightRatio, 1) ' 첫 번째 차원 (0 또는 1)
        For c = LBound(weightRatio, 2) To UBound(weightRatio, 2) ' 두 번째 차원 (0부터 19)
            weightRatio(r, c) = Empty
        Next c
    Next r
    ' --- (초기화 끝) ---

    '반영비율을 어레이에 저장하기 위한 표시번호 초기화
    arrayIndex = 0
    


    For r = LBound(weightRatioSKK, 2) To UBound(weightRatioSKK, 2)
    
        weightRatioSKK(0, r) = Empty
    Next r

   For iCol = weightRatioStartColNo To weightRatioEndColNo ' AD(30)~AZ(52)
        If examYearClassification = "2025학년도 정시" And _
            (targetUniversity = "성균관대학교" Or _
             targetUniversity = "인하대학교" _
            ) Then
        
            '성균관 대학교/ 인하대학교 특이한 비율을 할당하기 위함.
            position = 0
        
            On Error Resume Next
            
            position = WorksheetFunction.Find("(", wsJeongsi.Cells(jeongsiFoundRow, iCol).Value)
            On Error GoTo 0
            
            If position > 0 Then
            
                cellValue = CInt(Left(wsJeongsi.Cells(jeongsiFoundRow, iCol).Value, 2))
                weightRatio(0, arrayIndex) = cellValue
                weightRatio(1, arrayIndex) = Empty
                
                weightRatioSKK(0, arrayIndex) = CInt(Mid(wsJeongsi.Cells(jeongsiFoundRow, iCol).Value, position + 1, 2))
            Else
                
                cellValue = wsJeongsi.Cells(jeongsiFoundRow, iCol).Value
                weightRatio(0, arrayIndex) = cellValue
                weightRatioSKK(0, arrayIndex) = cellValue
                
                weightRatio(1, arrayIndex) = Empty
            End If
            
        Else
        
            cellValue = wsJeongsi.Cells(jeongsiFoundRow, iCol).Value
            weightRatio(0, arrayIndex) = cellValue
            weightRatio(1, arrayIndex) = Empty
            
        End If
        
        If Not IsError(cellValue) And IsNumeric(cellValue) And cellValue <> 0 Then
            Select Case iCol
                
                Case 31 ' 국어
                    weightRatio(1, arrayIndex) = koreanScore
                
                Case 32 ' 수학
                    weightRatio(1, arrayIndex) = mathScore
                
                Case 33 ' 영어
                    weightRatio(1, arrayIndex) = englishScore
                
                Case 34 ' 탐구1
                    weightRatio(1, arrayIndex) = WorksheetFunction.Max(inquiry1Score, inquiry2Score)
                
                Case 35 ' 탐구2 (탐구1, 탐구2 합산 또는 평균)
                
                    If targetUniversity = "가천대학교" Or _
                        targetUniversity = "강원대학교_춘천" Or _
                        targetUniversity = "단국대학교_죽전" Or _
                        targetUniversity = "단국대학교_천안" Or _
                        targetUniversity = "덕성여자대학교" Or _
                        targetUniversity = "동국대학교_경주" Or _
                        targetUniversity = "동덕여자대학교" Or _
                        targetUniversity = "삼육대학교" Or _
                        targetUniversity = "서경대학교" Or _
                        targetUniversity = "성신여자대학교" Or _
                        targetUniversity = "을지대학교" Or _
                        targetUniversity = "인하대학교" Or _
                        targetUniversity = "전북대학교" Or _
                        targetUniversity = "총신대학교" Or _
                        targetUniversity = "경상국립대학교" Or _
                        targetUniversity = "국립목포해양대학교" Or _
                        targetUniversity = "한국교원대학교" _
                    Then
                    
                        weightRatio(1, arrayIndex) = WorksheetFunction.Average(inquiry1Score, inquiry2Score)
                    Else
                        weightRatio(1, arrayIndex) = inquiry1Score + inquiry2Score
                    
                    End If
                    
                
                Case 36 ' 한국사
                    
                    weightRatio(1, arrayIndex) = koreanHistoryScore
                
                Case 37 ' 국수MAX (국어와 수학 중 큰 값)
                   
                    weightRatio(1, arrayIndex) = WorksheetFunction.Max(koreanScore, mathScore)
                
                Case 38 ' 국수MIN (국어와 수학 중 작은 값)
                    weightRatio(1, arrayIndex) = WorksheetFunction.Min(koreanScore, mathScore)
                
                Case 39 ' 국수탐MAX (국어, 수학, 탐구 중 큰 값, 탐구는 큰 탐구 점수)
                
                    If targetUniversity = "홍익대학교" Then
                        weightRatio(1, arrayIndex) = WorksheetFunction.Max(koreanScore, mathScore, WorksheetFunction.Sum(inquiry1Score, inquiry2Score))
                        
                    ElseIf targetUniversity = "서경대학교" Then
                    
                        weightRatio(1, arrayIndex) = WorksheetFunction.Max(koreanScore, mathScore, WorksheetFunction.Average(inquiry1Score, inquiry2Score))
                    
                    Else
                    
                        weightRatio(1, arrayIndex) = WorksheetFunction.Max(koreanScore, mathScore, WorksheetFunction.Max(inquiry1Score, inquiry2Score))
                    End If
                Case 40 ' 국수탐MID (중간 값)
                    ReDim arrMaxMidMin(0 To 2)
                    
                    For i = 0 To UBound(arrMaxMidMin)
                    
                        arrMaxMidMin(i) = Empty
                        
                    Next i
                    
                    arrMaxMidMin(0) = koreanScore
                    arrMaxMidMin(1) = mathScore
                    
                    If targetUniversity = "홍익대학교" Then
                    
                        arrMaxMidMin(2) = WorksheetFunction.Sum(inquiry1Score, inquiry2Score)
                        
                    ElseIf targetUniversity = "서경대학교" Then
                    
                        arrMaxMidMin(2) = WorksheetFunction.Average(inquiry1Score, inquiry2Score)
                    Else
                    
                        arrMaxMidMin(2) = WorksheetFunction.Max(inquiry1Score, inquiry2Score)
                    
                    End If
                    
                    weightRatio(1, arrayIndex) = WorksheetFunction.Large(arrMaxMidMin, 2)
                
                Case 41 ' 국수탐MIN (최소 값)
                    If targetUniversity = "홍익대학교" Then
                    
                        weightRatio(1, arrayIndex) = WorksheetFunction.Min(koreanScore, mathScore, WorksheetFunction.Sum(inquiry1Score, inquiry2Score))
                    ElseIf targetUniversity = "서경대학교" Then
                    
                        weightRatio(1, arrayIndex) = WorksheetFunction.Min(koreanScore, mathScore, WorksheetFunction.Average(inquiry1Score, inquiry2Score))
                    Else
                        weightRatio(1, arrayIndex) = WorksheetFunction.Min(koreanScore, mathScore, WorksheetFunction.Max(inquiry1Score, inquiry2Score))
                
                    End If
                Case 42 ' 국수영탐MAX (국어, 수학, 영어, 탐구 중 최대점)
'                    weightRatio(1, arrayIndex) = WorksheetFunction.Max(KoreanScore, MathScore, englishScore, WorksheetFunction.Max(inquiry1Score, inquiry2Score))
                    ReDim arrMaxMidMin(0 To 3)
                    
                    For i = 0 To UBound(arrMaxMidMin)
                    
                        arrMaxMidMin(i) = Empty
                        
                    Next i
                    
                    arrMaxMidMin(0) = koreanScore
                    arrMaxMidMin(1) = mathScore
                    arrMaxMidMin(2) = englishScore
                    
                    If targetUniversity = "삼육대학교" Then
                        If koreanHistoryScore > WorksheetFunction.Min(inquiry1Score, inquiry2Score) Then
                            
                            arrMaxMidMin(3) = WorksheetFunction.Average(koreanHistoryScore, WorksheetFunction.Max(inquiry1Score, inquiry2Score))

                        Else

                            arrMaxMidMin(3) = WorksheetFunction.Average(inquiry1Score, inquiry2Score)
                        End If
                    
                    ElseIf targetUniversity = "을지대학교" Then
                        
                        arrMaxMidMin(3) = WorksheetFunction.Average(inquiry1Score, inquiry2Score)
                    
                    Else
                        arrMaxMidMin(3) = WorksheetFunction.Max(inquiry1Score, inquiry2Score)
                    End If
                    
                    
                    weightRatio(1, arrayIndex) = WorksheetFunction.Large(arrMaxMidMin, 1)
                    
                Case 43 ' 국수영탐MID1 (4개 중 첫 번째 중간값 - 필요 시 다른 중간값 로직으로 변경 가능)
                    ' 간단히 4개 중 2번째 값 반환용 함수 미구현. 기본으로 영어 점수 반영 중간값 예제 처리
                    ReDim arrMaxMidMin(0 To 3)
                    
                    For i = 0 To UBound(arrMaxMidMin)
                    
                        arrMaxMidMin(i) = Empty
                        
                    Next i
                    
                    arrMaxMidMin(0) = koreanScore
                    arrMaxMidMin(1) = mathScore
                    arrMaxMidMin(2) = englishScore
                    If targetUniversity = "삼육대학교" Then
                        If koreanHistoryScore > WorksheetFunction.Min(inquiry1Score, inquiry2Score) Then
                            
                            arrMaxMidMin(3) = WorksheetFunction.Average(koreanHistoryScore, WorksheetFunction.Max(inquiry1Score, inquiry2Score))

                        Else

                            arrMaxMidMin(3) = WorksheetFunction.Average(inquiry1Score, inquiry2Score)
                        End If
                    
                    ElseIf targetUniversity = "을지대학교" Then
                        
                        arrMaxMidMin(3) = WorksheetFunction.Average(inquiry1Score, inquiry2Score)
                    
                    Else
                        arrMaxMidMin(3) = WorksheetFunction.Max(inquiry1Score, inquiry2Score)
                    End If
                    weightRatio(1, arrayIndex) = WorksheetFunction.Large(arrMaxMidMin, 2)
                
                Case 44 ' 국수영탐MID2 (4개 중 두 번째 중간값 - 필요 시 구현)
                    ' 필요에 따라 다른 중간값 산출 함수 추가 가능
                    ReDim arrMaxMidMin(0 To 3)
                    
                    For i = 0 To UBound(arrMaxMidMin)
                    
                        arrMaxMidMin(i) = Empty
                        
                    Next i
                    
                    arrMaxMidMin(0) = koreanScore
                    arrMaxMidMin(1) = mathScore
                    arrMaxMidMin(2) = englishScore
                    If targetUniversity = "삼육대학교" Then
                        If koreanHistoryScore > WorksheetFunction.Min(inquiry1Score, inquiry2Score) Then
                            
                            arrMaxMidMin(3) = WorksheetFunction.Average(koreanHistoryScore, WorksheetFunction.Max(inquiry1Score, inquiry2Score))

                        Else

                            arrMaxMidMin(3) = WorksheetFunction.Average(inquiry1Score, inquiry2Score)
                        End If
                    
                    ElseIf targetUniversity = "을지대학교" Then
                        
                        arrMaxMidMin(3) = WorksheetFunction.Average(inquiry1Score, inquiry2Score)
                    
                    Else
                        arrMaxMidMin(3) = WorksheetFunction.Max(inquiry1Score, inquiry2Score)
                    End If
                    weightRatio(1, arrayIndex) = WorksheetFunction.Large(arrMaxMidMin, 3)
                    
                Case 45 ' 국수영탐MIN (4개 중 최소값)
                    ReDim arrMaxMidMin(0 To 3)
                    
                    For i = 0 To UBound(arrMaxMidMin)
                    
                        arrMaxMidMin(i) = Empty
                        
                    Next i
                    
                    arrMaxMidMin(0) = koreanScore
                    arrMaxMidMin(1) = mathScore
                    arrMaxMidMin(2) = englishScore
                    If targetUniversity = "삼육대학교" Then
                        If koreanHistoryScore > WorksheetFunction.Min(inquiry1Score, inquiry2Score) Then
                            
                            arrMaxMidMin(3) = WorksheetFunction.Average(koreanHistoryScore, WorksheetFunction.Max(inquiry1Score, inquiry2Score))

                        Else

                            arrMaxMidMin(3) = WorksheetFunction.Average(inquiry1Score, inquiry2Score)
                        End If
                    
                    ElseIf targetUniversity = "을지대학교" Then
                        
                        arrMaxMidMin(3) = WorksheetFunction.Average(inquiry1Score, inquiry2Score)
                    
                    Else
                        arrMaxMidMin(3) = WorksheetFunction.Max(inquiry1Score, inquiry2Score)
                    End If
                    weightRatio(1, arrayIndex) = WorksheetFunction.Large(arrMaxMidMin, 4)
                
                Case 46 ' 수탐MAX (수학과 탐구 중 최대값)
                    If targetUniversity = "을지대학교" Then
                    
                        weightRatio(1, arrayIndex) = WorksheetFunction.Max(mathScore, WorksheetFunction.Average(inquiry1Score, inquiry2Score))
                    Else
                    
                        weightRatio(1, arrayIndex) = WorksheetFunction.Max(mathScore, WorksheetFunction.Max(inquiry1Score, inquiry2Score))
                    End If
                    
                    
                Case 47 ' 국탐MAX (국어와 탐구 최대값)
                    If targetUniversity = "을지대학교" Then
                        weightRatio(1, arrayIndex) = WorksheetFunction.Max(koreanScore, WorksheetFunction.Average(inquiry1Score, inquiry2Score))
                    Else
                        weightRatio(1, arrayIndex) = WorksheetFunction.Max(koreanScore, WorksheetFunction.Max(inquiry1Score, inquiry2Score))
                    End If
                    
                
                Case 48 ' 국탐MIN (국어와 탐구 최소값)
                    weightRatio(1, arrayIndex) = WorksheetFunction.Min(koreanScore, WorksheetFunction.Max(inquiry1Score, inquiry2Score))
                
                Case 49 ' 국영탐MAX (국어, 영어, 탐구 최대값)
                    ReDim arrMaxMidMin(0 To 2)
                    
                    For i = 0 To UBound(arrMaxMidMin)
                    
                        arrMaxMidMin(i) = Empty
                        
                    Next i
                    
                    arrMaxMidMin(0) = koreanScore
                    arrMaxMidMin(1) = englishScore
                    arrMaxMidMin(2) = WorksheetFunction.Max(inquiry1Score, inquiry2Score)
                    weightRatio(1, arrayIndex) = WorksheetFunction.Large(arrMaxMidMin, 1)
                
                Case 50 ' 국영탐MID (국어, 영어, 탐구 중간값)
                    ReDim arrMaxMidMin(0 To 2)
                    
                    For i = 0 To UBound(arrMaxMidMin)
                    
                        arrMaxMidMin(i) = Empty
                        
                    Next i
                    
                    arrMaxMidMin(0) = koreanScore
                    arrMaxMidMin(1) = englishScore
                    arrMaxMidMin(2) = WorksheetFunction.Max(inquiry1Score, inquiry2Score)
                    weightRatio(1, arrayIndex) = WorksheetFunction.Large(arrMaxMidMin, 2)
                
                Case 51 '수영탐MAX
                    ReDim arrMaxMidMin(0 To 2)
                    
                    For i = 0 To UBound(arrMaxMidMin)
                    
                        arrMaxMidMin(i) = Empty
                        
                    Next i
                    
                    arrMaxMidMin(0) = mathScore
                    arrMaxMidMin(1) = englishScore
                    arrMaxMidMin(2) = WorksheetFunction.Max(inquiry1Score, inquiry2Score)
                    weightRatio(1, arrayIndex) = WorksheetFunction.Large(arrMaxMidMin, 1)
                
                Case 52 '수영탐MID
                
                    ReDim arrMaxMidMin(0 To 2)
                    
                    For i = 0 To UBound(arrMaxMidMin)
                    
                        arrMaxMidMin(i) = Empty
                        
                    Next i
                    
                    arrMaxMidMin(0) = mathScore
                    arrMaxMidMin(1) = englishScore
                    arrMaxMidMin(2) = WorksheetFunction.Max(inquiry1Score, inquiry2Score)
                    weightRatio(1, arrayIndex) = WorksheetFunction.Large(arrMaxMidMin, 2)
                
                
                
                Case 53 '수영MAX
                
                    weightRatio(1, arrayIndex) = WorksheetFunction.Max(mathScore, englishScore)
                
                Case 54 '국영MAX
                    weightRatio(1, arrayIndex) = WorksheetFunction.Max(koreanScore, englishScore)
                
                Case 55 '국영수MAX
                    weightRatio(1, arrayIndex) = WorksheetFunction.Max(koreanScore, englishScore, mathScore)
                    
                Case 56 '국영수MID
                    ReDim arrMaxMidMin(0 To 2)
                    
                    For i = 0 To UBound(arrMaxMidMin)
                    
                        arrMaxMidMin(i) = Empty
                        
                    Next i
                    
                    arrMaxMidMin(0) = koreanScore
                    arrMaxMidMin(1) = englishScore
                    arrMaxMidMin(2) = mathScore
                    weightRatio(1, arrayIndex) = WorksheetFunction.Large(arrMaxMidMin, 2)
                
                Case 57 '국영수MIN
                    ReDim arrMaxMidMin(0 To 2)
                    
                    For i = 0 To UBound(arrMaxMidMin)
                    
                        arrMaxMidMin(i) = Empty
                        
                    Next i
                    
                    arrMaxMidMin(0) = koreanScore
                    arrMaxMidMin(1) = englishScore
                    arrMaxMidMin(2) = mathScore
                    weightRatio(1, arrayIndex) = WorksheetFunction.Large(arrMaxMidMin, 3)
                
                Case Else
                    weightRatio(1, arrayIndex) = cellValue ' 필요시 기본값
            End Select
        End If
        
        arrayIndex = arrayIndex + 1
    Next iCol

    ' --- [결과 확인 (선택 사항)] ---
    Dim j As Long
    For j = 0 To UBound(weightRatio, 2)
        Debug.Print "weightRatio[" & j & "] - 찾은 값: " & IIf(IsEmpty(weightRatio(0, j)), "[없음]", weightRatio(0, j)) & _
                    ", 연계된 값: " & IIf(IsEmpty(weightRatio(1, j)), "[없음]", weightRatio(1, j))
    Next j

'    'MsgBox "AD열부터 20개 열에 대한 검색을 완료하고, 모듈 수준 'weightRatio' 배열에 결과를 저장했습니다."

End Sub



'목표대학, 학과에 대한 반영 비율 weightRatio[], 가산점 addBonus[]에 저장됨. 이를 활용
Sub TargetUniversityScore()
    
    Dim i As Integer
    Dim lastRowStdScoreMax As Integer
    Dim totalSum1 As Long, totalSum2 As Long
    Dim inquirySatStdScoreMax As Integer
    
    totalSum1 = 0
    totalSum2 = 0
    '============================================================================
    '수능 표준점수의 과목별 최고점을 저장해 두자.
    '영어의 경우에는 따로 변환점수의 최고점을 저정해 두자
'    If bSubjectSatStdScoreMaxFlag = 0 And _
'        ( _
'            targetUniversity = "건국대학교_글로컬" Or _
'            targetUniversity = "단국대학교_천안" Or _
'            targetUniversity = "서울과학기술대학교" Or _
'            targetUniversity = "서울시립대학교" Or _
'            targetUniversity = "숙명여자대학교" Or _
'            targetUniversity = "숭실대학교" Or _
'            targetUniversity = "아주대학교" Or _
'            targetUniversity = "이화여자대학교" Or _
'            targetUniversity = "한국외국어대학교" Or _
'            targetUniversity = "한국외국어대학교_글로벌캠퍼스" Or _
'            targetUniversity = "한양대학교" Or _
'            targetUniversity = "한양대학교_에리카" Or _
'            targetUniversity = "충북대학교" Or _
'            targetUniversity = "전남대학교" Or _
'            targetUniversity = "경상국립대학교" Or _
'            targetUniversity = "인하대학교" Or _
'            targetUniversity = "고려대학교_세종" _
'        ) Then
'
'        lastRowStdScoreMax = wsSatStdScoreMax.Cells(wsSatStdScoreMax.Rows.Count, "A").End(xlUp).Row
'
'        For i = 1 To lastRowStdScoreMax
'
'            If wsSatStdScoreMax.Cells(i, "A").Value = examYearClassification Then
'
'                If wsSatStdScoreMax.Cells(i, "B").Value = "국어" Then
'                    subjectSatStdScoreMax(0) = wsSatStdScoreMax.Cells(i, "C").Value
'
'
'                ElseIf wsSatStdScoreMax.Cells(i, "B").Value = "수학" Then
'                    subjectSatStdScoreMax(1) = wsSatStdScoreMax.Cells(i, "C").Value
'
'                ElseIf wsSatStdScoreMax.Cells(i, "B").Value = inquiry1SubjectName Then
'
'                    '아주대학교는 변환표준점수 계산할 때 이미 저장해 둠
'                    If targetUniversity <> "아주대학교" Then
'                        subjectSatStdScoreMax(3) = wsSatStdScoreMax.Cells(i, "C").Value
'                    End If
'
'                ElseIf wsSatStdScoreMax.Cells(i, "B").Value = inquiry2SubjectName Then
'
'                    '아주대학교는 변환표준점수 계산할 때 이미 저장해 둠
'                    If targetUniversity <> "아주대학교" Then
'                        subjectSatStdScoreMax(4) = wsSatStdScoreMax.Cells(i, "C").Value
'                    End If
'                End If
'
'            End If
'        Next i
'
'
'        bSubjectSatStdScoreMaxFlag = bSubjectSatStdScoreMaxFlag + 1 '대학별 검색에서 한 번만 작업을 하도록 Flag를 넣음.
'
'    End If
'
'    '영어 만점의 점수가 학과별로 바뀌는 경우가 있어서 분리함.
'    If wsJeongsi.Cells(jeongsiFoundRow, colJeongsiEnglishGrade1).Value <> "" Then
'
'        subjectSatStdScoreMax(2) = wsJeongsi.Cells(jeongsiFoundRow, colJeongsiEnglishGrade1).Value
'
'    End If
'
'    '==========================================================================
    
    Select Case targetUniversity
    
        Case "가천대학교"
            ' 가천대학교에 대한 처리 작성
            '======================================
            '2024학년도, 2025학년도 반영 방법 일치
            '======================================
            For i = 0 To weightRatioColCount - 1
                totalSum = totalSum + weightRatio(0, i) * weightRatio(1, i) * 0.01
            Next i
            
            totalSum = totalSum + koreanHistoryScore
            
        Case "건국대학교" '건국대학교_글로컬 다시 검토필요
            ' 건국대학교에 대한 처리 작성
            konkuk
            
            
        Case "건국대학교_글로컬"
        
            konkuk_glocal
            
        Case "경기대학교"
            ' 경기대학교에 대한 처리 작성
            '======================================
            '2024학년도, 2025학년도 반영 방법 일치
            '======================================

            For i = 0 To weightRatioColCount - 1
                totalSum = totalSum + weightRatio(0, i) * weightRatio(1, i) * 0.01
            Next i
        
'            totalSum = totalSum * wsJeongsi.Cells(jeongsiFoundRow, colJeongsiMaxScore).Value / 200
'==================================================================================
            '유아교육과와 체육학고의 0.95, 0.7을 곱한 결과가 대학 발표의 성적인가?
            '인적성 결과도 포함한 값인지 확인해야함.
            '인적성도 포함한 값이면 유아교육에는 0.95를 한 값에 5점을 추가해야 할 듯한다.
            '이 문의를 하고 곱한 값만 발표한 것이면 주석을 풀자
'            If targetMajor = "유아교육과" Then
'
'                totalSum = totalSum * 0.95
'
'            ElseIf targetMajor = "체육학과" Then
'
'                totalSum = totalSum * 0.7
'            Else
'                totalSum = totalSum + koreanHistoryScore
'            End If
'            '이 주석을 풀면 아래의 totalsum 처리를 주석처리한다.
'==================================================================================
           
            
            totalSum = totalSum + koreanHistoryScore
            
        Case "경북대학교"
            Kyungpook
            
        Case "경희대학교"
            ' 경희대학교에 대한 처리 작성
            '2024학년도와 2025학년도 처리방식 동일
                If weightRatio(0, 0) = 50 Then

                    For i = 0 To weightRatioColCount - 1

                        If i = 3 Then
                            totalSum = totalSum + weightRatio(0, i) * (weightRatio(1, i) + 100) * 0.01
                        Else
                            totalSum = totalSum + weightRatio(0, i) * weightRatio(1, i) * 0.01

                        End If
                    Next i

                Else
                    For i = 0 To weightRatioColCount - 1
                            totalSum = totalSum + weightRatio(0, i) * weightRatio(1, i) * 0.01

                    Next i
                End If
                
            totalSum = (totalSum + koreanHistoryScore) / 2 * wsJeongsi.Cells(jeongsiFoundRow, colJeongsiMaxScore) / 100
        
        
        Case "고려대학교"
            ' 고려대학교에 대한 처리 작성
            ' 2024학년도와 2025학년도의 처리방식이 동일하다.
            
            
            For i = 0 To weightRatioColCount - 1

                '학생의 점수 영역 반영점수 200이면 1을 곱하고, 240이면 1.2를 160이면 0.8을 표준점수에 곱한다.
                totalSum = totalSum + weightRatio(1, i) * weightRatio(0, i) / 200
                totalSum1 = totalSum1 + weightRatio(0, i) '반영 점수 ex. 국어 200, 수학 240 탐구 160 등...


            Next i
            totalSum = totalSum / totalSum1 * wsJeongsi.Cells(jeongsiFoundRow, colJeongsiMaxScore) + koreanHistoryScore + englishScore
            
        Case "고려대학교_세종"
        
            '2024학년도 변환표준점수 표가 없다.....
            '어디서 구해야하나??????
        
            If examYearClassification = "2024학년도 정시" Then
            
                inquirySatStdScoreMax = 71.75
                
            ElseIf examYearClassification = "2025학년도 정시" Then
            
                inquirySatStdScoreMax = 70.12
            
            
            End If
            
            For i = 0 To weightRatioColCount - 1

                If IsNumeric(weightRatio(0, i)) And weightRatio(0, i) > 0 Then
                    If i = 4 Then
                    
                        totalSum = totalSum + weightRatio(1, i) * weightRatio(0, i) / 100 '탐구과목 변환 표준점수의 합
                        totalSum1 = totalSum1 + inquirySatStdScoreMax * weightRatio(0, i) * 2 / 100 '탐구 변환표준점수 최고점의 두배...
                    
                    Else
                    
                        totalSum = totalSum + weightRatio(1, i) * weightRatio(0, i) / 100 '학생의 점수
                        totalSum1 = totalSum1 + subjectSatStdScoreMax(i) * weightRatio(0, i) / 100 '반영 점수 ex. 국어 200, 수학 240 탐구 160 등...
                
                    End If
                End If


            Next i
        
            totalSum = totalSum / totalSum1 * wsJeongsi.Cells(jeongsiFoundRow, colJeongsiMaxScore)
        
        
        Case "광운대학교"
            ' 광운대학교에 대한 처리 작성
            
            Kwangwoon
            
            
        Case "국립목포해양대학교"
            MMU
        
        Case "국민대학교"
            ' 국민대학교에 대한 처리 작성
            Kookmin
            
        Case "단국대학교_죽전"
            Dankook
        
        Case "단국대학교_천안"
            Dankook
        
            
        Case "덕성여자대학교"
            ' 덕성여자대학교에 대한 처리 작성
            '2024학년와 2025학년도 반영방법 동일
            
            For i = 0 To weightRatioColCount - 1
                totalSum = totalSum + weightRatio(0, i) * weightRatio(1, i) * 0.01
            Next i
        
            totalSum = totalSum * wsJeongsi.Cells(jeongsiFoundRow, colJeongsiMaxScore).Value / 100 + koreanHistoryScore
            
            
        
        Case "동국대학교", "동국대학교_경주"
            ' 동국대학교에 대한 처리 작성
            Dongguk
  
        Case "동덕여자대학교"
            ' 동덕여자대학교에 대한 처리 작성
            Dongduk
            
        Case "명지대학교"
        
            Myongji
            
            ' 명지대학교에 대한 처리 작성
            
            
        Case "부산대학교"
            Pusan
        
        Case "삼육대학교"
            ' 삼육대학교에 대한 처리 작성
            
            Sahmyook
            
            
        
        Case "상명대학교", "상명대학교_천안"  ' 상명대학교_천안은 다시 검토필요
            ' 상명대학교에 대한 처리 작성
            
            Sangmyung
        
        Case "서강대학교"
            ' 서강대학교에 대한 처리 작성
            Sogang
            
        Case "서경대학교"
            For i = 0 To weightRatioColCount - 1
                totalSum = totalSum + weightRatio(0, i) * weightRatio(1, i) * 0.01
            Next i
        
            totalSum = totalSum * wsJeongsi.Cells(jeongsiFoundRow, colJeongsiMaxScore).Value / 100 + koreanHistoryScore
            
        Case "서울과학기술대학교"
            ' 서울과학기술대학교에 대한 처리 작성
            SeoulTech
            
        Case "서울대학교"
            
            ' 서울대학교에 대한 처리 작성
            SNU
            
        Case "서울교육대학교"
            SNUE
        
        
        Case "서울시립대학교"
            ' 서울시립대학교에 대한 처리 작성
        
            UOS
            
            
        Case "서울여자대학교"
            ' 서울여자대학교에 대한 처리 작성
        
            SWU
            
        Case "서울한영대학교"
            ' 서울한영대학교에 대한 처리 작성
            '2024학년도와 2025학년도가 동일
            
            For i = 0 To weightRatioColCount - 1
                totalSum = totalSum + weightRatio(0, i) * weightRatio(1, i) * 0.01
            Next i
        
            totalSum = totalSum * wsJeongsi.Cells(jeongsiFoundRow, colJeongsiMaxScore).Value / 100
            
        Case "성공회대학교"
            ' 성공회대학교에 대한 처리 작성
            Sungkonghoe
            
            
        Case "성균관대학교"
            ' 성균관대학교에 대한 처리 작성
            
             SKK
               
        Case "성신여자대학교"
            ' 성신여자대학교에 대한 처리 작성
            Sungshin
            
        Case "세종대학교"
            ' 세종대학교에 대한 처리 작성
            
            Sejong
            
        Case "숙명여자대학교"
            ' 숙명여자대학교에 대한 처리 작성
            Sookmyung
            
        Case "숭실대학교"
            ' 숭실대학교에 대한 처리 작성
            Soongsil
            
        Case "아주대학교"
        
            Ajou
        Case "연세대학교"
            ' 연세대학교에 대한 처리 작성
            Yonsei
        Case "연세대학교_미래"
        
            Yonsei_Mirae
            
        Case "을지대학교"
            ' 을지대학교에 대한 처리 작성
            Eulji
                        
        Case "이화여자대학교"
            ' 이화여자대학교에 대한 처리 작성
            
            Ewha
            
        Case "중앙대학교"
            ' 중앙대학교에 대한 처리 작성
            CAU
            
        Case "총신대학교"
        
            ' 총신대학교에 대한 처리 작성
            '===================================================
            '2024학년도와 2025학년도 반영 방법 변경사항으로 분리
            '===================================================
            For i = 0 To weightRatioColCount - 1
                totalSum = totalSum + weightRatio(0, i) * weightRatio(1, i) * 0.01 '100점 만점으로 계산
            Next i
        
            
            totalSum = totalSum * wsJeongsi.Cells(jeongsiFoundRow, colJeongsiMaxScore).Value / 100 + koreanHistoryScore
            
        Case "가톨릭대학교"
            ' 가톨릭대학교에 대한 처리 작성
            Catholic
            
        Case "한국외국어대학교", "한국외국어대학교_글로벌캠퍼스"
            ' 한국외국어대학교에 대한 처리 작성
            HUFS
            
        Case "한남대학교"
            '===================================================
            '2024학년도와 2025학년도 반영 방법 변경사항으로 분리
            '===================================================
            For i = 0 To weightRatioColCount - 1
                totalSum = totalSum + weightRatio(1, i)   '300점 만점으로 계산
            Next i
            
            
            totalSum = totalSum * wsJeongsi.Cells(jeongsiFoundRow, colJeongsiMaxScore).Value / 300 + koreanHistoryScore
            
        Case "한성대학교"
            ' 한성대학교에 대한 처리 작성
            '===================================================
            '2024학년도와 2025학년도 반영 방법 변경사항으로 분리
            '===================================================
            For i = 0 To weightRatioColCount - 1
            
                If i = 2 And weightRatio(1, i) = 100 Then
                
                    totalSum = totalSum + weightRatio(0, i) * 2 * weightRatio(1, i) * 0.01 '영어1등급은 2배의 점수를 줌.
                    
                Else
                
                    totalSum = totalSum + weightRatio(0, i) * weightRatio(1, i) * 0.01 '100점 만점으로 계산
                End If
            Next i
            
            totalSum = totalSum * wsJeongsi.Cells(jeongsiFoundRow, colJeongsiMaxScore).Value / 100 + koreanHistoryScore
            
            
        Case "한양대학교"
            ' 한양대학교에 대한 처리 작성
            Hanyang
        
        Case "한양대학교_에리카"
        
            HanyangErica
            
        Case "홍익대학교", "홍익대학교_세종"
            ' 홍익대학교 서울캠퍼스에 대한 처리 작성
            '===================================================
            '2024학년도와 2025학년도 반영 방법 변경사항으로 분리
            '===================================================
            For i = 0 To weightRatioColCount - 1
                totalSum = totalSum + weightRatio(0, i) * weightRatio(1, i) * 0.01
            Next i
            totalSum = totalSum + koreanHistoryScore
        
        
        Case "충남대학교"
            Chungnam
            
        
        Case "충북대학교"
        
            Chungbuk
        
        Case "전남대학교"
            Chonnam
        
        Case "전북대학교"
            Jeonbuk
        
        
        
        Case "경상국립대학교"
            Gyeongsang
        
        Case "강원대학교_춘천"
            Kangwon_Chuncheon
        
        Case "인천대학교"
            Incheon
        
        Case "인하대학교"
            Inha
        
        Case "한국교원대학교"
            KNUE
            
        Case Else
            ' 해당하지 않는 대학 처리 또는 메시지 출력
            MsgBox "지원하는 대학 이름이 목록에 없습니다: " & targetUniversity, vbExclamation
    End Select
    
    
End Sub

Sub Catholic() '가톨릭대학교
    '===================================================
    '2024학년도와 2025학년도 반영 방법 변경사항으로 분리
    '2025와 2026은 일치
    '===================================================
    
    For i = 0 To weightRatioColCount - 1
        totalSum = totalSum + weightRatio(0, i) * weightRatio(1, i) * 0.01
    Next i
        
    If examYearClassification = "2024학년도 정시" Then
    
        If targetMajor = "의예과" Or _
           targetMajor = "약학과" Or _
           targetMajor = "간호학과(인문)" Or _
           targetMajor = "간호학과(자연)" Then
           
            totalSum = totalSum * 5 + englishScore + koreanHistoryScore
           
        Else
        
                totalSum = totalSum * 10 + koreanHistoryScore
        
        End If
    
    ElseIf examYearClassification = "2025학년도 정시" Or _
            examYearClassification = "2026학년도 정시" Then
        If targetMajor = "의예과" Then
        
            totalSum = (totalSum * 5 + englishScore + koreanHistoryScore) * 0.95
            
        ElseIf targetMajor = "간호학과" Or _
           targetMajor = "약학과" Then
        
            totalSum = totalSum * 5 + englishScore + koreanHistoryScore
        Else
            If wsJeongsi.Cells(jeongsiFoundRow, "D") = "일반전형Ⅰ" Then
                totalSum = totalSum * 10 + koreanHistoryScore
            ElseIf wsJeongsi.Cells(jeongsiFoundRow, "D") = "일반전형Ⅱ" Then
                totalSum = (totalSum + englishScore) * 10 + koreanHistoryScore
            End If
        End If
    End If
    
    totalSum = Application.WorksheetFunction.Round(totalSum, 4) '지원자의 대학 환산점수
End Sub
Sub konkuk_glocal()
'======================================
'2024학년도, 2025학년도 반영 방법 일치
'======================================
    For i = 0 To weightRatioColCount - 1
        If i = 0 And IsNumeric(weightRatio(0, i)) And weightRatio(0, i) > 0 Then
        
           totalSum = totalSum + weightRatio(0, i) * weightRatio(1, i) / subjectSatStdScoreMax(0) * 0.01  '국어
           
           
           Debug.Print "국어 표점최고:" & subjectSatStdScoreMax(0)
           
        ElseIf i = 1 And IsNumeric(weightRatio(0, i)) And weightRatio(0, i) > 0 Then
        
           totalSum = totalSum + weightRatio(0, i) * weightRatio(1, i) / subjectSatStdScoreMax(1) * 0.01 '수학
           
           Debug.Print "수학 표점최고:" & subjectSatStdScoreMax(1)
           
           
           
        ElseIf i = 2 And IsNumeric(weightRatio(0, i)) And weightRatio(0, i) > 0 Then
        
               
            totalSum = totalSum + weightRatio(0, i) * weightRatio(1, i) / subjectSatStdScoreMax(2) * 0.01  '영어
            Debug.Print "영어 표점최고:" & subjectSatStdScoreMax(2)
           
           
        ElseIf i = 4 And IsNumeric(weightRatio(0, i)) And weightRatio(0, i) > 0 Then
        
            If Len(inquiry1SubjectName) > 0 Then
        
                inquiry1ScoreImsi = inquiry1Score / subjectSatStdScoreMax(3)
                
                Debug.Print inquiry1SubjectName & " 표점최고:" & subjectSatStdScoreMax(3)
                
            End If
            
            If Len(inquiry2SubjectName) > 0 Then
                inquiry2ScoreImsi = inquiry2Score / subjectSatStdScoreMax(4)
                
                Debug.Print inquiry2SubjectName & " 표점최고:" & subjectSatStdScoreMax(4)
            End If
            
            weightRatio(1, i) = WorksheetFunction.Average(inquiry1ScoreImsi, inquiry2ScoreImsi)
        
           totalSum = totalSum + weightRatio(0, i) * weightRatio(1, i) * 0.01
        Else
            totalSum = totalSum + weightRatio(0, i) * weightRatio(1, i) * 0.01
        
        End If
    Next i
    
    If InStr(targetMajor, "의예과") Then
    
        totalSum = totalSum * wsJeongsi.Cells(jeongsiFoundRow, colJeongsiMaxScore).Value
        
    End If
    

    totalSum = totalSum + koreanHistoryScore

End Sub
Sub konkuk()
'======================================
'2024학년도, 2025학년도 반영 방법 일치
'======================================
    
    For i = 0 To weightRatioColCount - 1
        totalSum = totalSum + weightRatio(0, i) * weightRatio(1, i) * 0.01
    Next i
    totalSum = totalSum * wsJeongsi.Cells(jeongsiFoundRow, colJeongsiMaxScore).Value / 200
    
    totalSum = Application.WorksheetFunction.Round(totalSum, 4) '지원자의 대학 환산점수
    
    totalSum = totalSum + koreanHistoryScore
End Sub
Sub Kwangwoon() '광운대학교
    '========================================
    '2024학년도와 2025학년도 변화가 없음.
    '========================================
    
    For i = 0 To weightRatioColCount - 1
        totalSum = totalSum + weightRatio(0, i) * weightRatio(1, i) * 0.01 '200점 만점으로 계산
    Next i
    
    totalSum = totalSum * wsJeongsi.Cells(jeongsiFoundRow, colJeongsiMaxScore).Value / 200 + koreanHistoryScore  '200점을  만점으로 변환한 후에 한국사 점수 가점
End Sub

Sub Kookmin() '국민대학교

    '===============================
    '2024학년도와 2025학년도 동일
    '================================
    Dim totalSum1 As Double
    
    totalSum1 = 0
    

    For i = 0 To weightRatioColCount - 1
        If i = 2 And weightRatio(0, i) > 0 Then
        
            totalSum = totalSum + 2 * weightRatio(0, i) * weightRatio(1, i)
        
        Else
            totalSum = totalSum + weightRatio(0, i) * weightRatio(1, i)
        End If
        
        totalSum1 = totalSum1 + weightRatio(0, i) '수능반영영역 반영영역점수의 합.
    Next i
    
    totalSum = totalSum / 200 * wsJeongsi.Cells(jeongsiFoundRow, colJeongsiMaxScore).Value / totalSum1
End Sub

Sub Dankook() '단국대학교, 단국대학교_천안

    Dim targetRange As Range
    Set targetRange = ThisWorkbook.Sheets("수능표준점수최고점").Range("B2:C20")
    
    
    For i = 0 To weightRatioColCount - 1
    
        If targetMajor = "의예과" Or _
            targetMajor = "치의예과" Or _
            targetMajor = "약학과" Then
            
            If i = 0 And IsNumeric(weightRatio(0, i)) And weightRatio(0, i) > 0 Then
            
               totalSum = totalSum + weightRatio(0, i) * weightRatio(1, i) / subjectSatStdScoreMax(0)  '국어
               
            ElseIf i = 1 And IsNumeric(weightRatio(0, i)) And weightRatio(0, i) > 0 Then
            
               totalSum = totalSum + weightRatio(0, i) * weightRatio(1, i) / subjectSatStdScoreMax(1)  '수학
            Else
            
                totalSum = totalSum + weightRatio(0, i) * weightRatio(1, i) * 0.01
            
            End If
        Else
            totalSum = totalSum + weightRatio(0, i) * weightRatio(1, i) * 0.01
        End If
    Next i
    
    totalSum = totalSum * wsJeongsi.Cells(jeongsiFoundRow, colJeongsiMaxScore).Value / 100 + koreanHistoryScore
            
      



End Sub

Sub Dongguk() '동국대학교, 동국대학교_경주

    For i = 0 To weightRatioColCount - 1
        totalSum = totalSum + weightRatio(0, i) * weightRatio(1, i) * 0.01 ' 200점 만점으로 계산
    Next i
    
    If targetUniversity = "동국대학교" Then
    
        totalSum = totalSum / 200 * wsJeongsi.Cells(jeongsiFoundRow, colJeongsiMaxScore).Value + koreanHistoryScore '200점 만점을 1000점 만점으로 바꾸고 한국사점수 가산
        
    Else
        
        '동국대학교 경주의 2024학년도 반영 법이다. 동국대학교_경주의 2025학년도를 추가하면 다시 한 번 봐야한다.
        
        totalSum = totalSum / 100 * wsJeongsi.Cells(jeongsiFoundRow, colJeongsiMaxScore).Value + koreanHistoryScore '100점 만점을 1000점 만점으로 바꾸고 한국사점수 가산
    End If
            
      
End Sub

Sub Dongduk() '동덕여자대학교


    For i = 0 To weightRatioColCount - 1
    
        totalSum = totalSum + weightRatio(0, i) * weightRatio(1, i) * 0.01
                
    Next i

    totalSum = totalSum * wsJeongsi.Cells(jeongsiFoundRow, colJeongsiMaxScore).Value / 100 + koreanHistoryScore




End Sub

Sub Myongji() '명지대학교

    '=================================================
    '2024학년과 2025학년도 반영 방법 동일
    '=================================================
    
    For i = 0 To weightRatioColCount - 1
        totalSum = totalSum + weightRatio(0, i) * weightRatio(1, i) * 0.01 '100점 만점으로 계산
    Next i

    totalSum = totalSum * wsJeongsi.Cells(jeongsiFoundRow, colJeongsiMaxScore).Value / 100 + koreanHistoryScore  '100점 만점을 1000점 만점으로 변환 후 한국사 가산
            

End Sub
Sub Sahmyook() '삼육대학교

    '=================================================
    '2024학년과 2025학년도 반영 방법 동일
    '=================================================
   
    For i = 0 To weightRatioColCount - 1
    
        totalSum = totalSum + weightRatio(0, i) * weightRatio(1, i) * 0.01   '100점 만점으로 계산됨
    
    Next i
    
    totalSum = totalSum * wsJeongsi.Cells(jeongsiFoundRow, colJeongsiMaxScore).Value / 100  '100점 만점을 1000점 만점으로 환산
    

End Sub


Sub Sangmyung() '상명대학교
    '=================================================
    '2024학년과 2025학년도 반영 방법 동일
    '=================================================
     
     For i = 0 To weightRatioColCount - 1
         totalSum = totalSum + weightRatio(0, i) * weightRatio(1, i) * 0.01  '100점 만점으로 계산
         
     Next i
    
     totalSum = totalSum * wsJeongsi.Cells(jeongsiFoundRow, colJeongsiMaxScore).Value / 100 + koreanHistoryScore
            
End Sub


Sub Sogang() '서강대학교
    '=================================================
    '2024학년과 2025학년도 반영 방법 동일
    '=================================================
    For i = 0 To weightRatioColCount - 1
    
       If i = 4 Then
           totalSum = totalSum + weightRatio(0, i) * (weightRatio(1, i) + bonusCase8)
       
       Else
           totalSum = totalSum + weightRatio(0, i) * weightRatio(1, i)
       End If
    Next i
     
    totalSum = Application.WorksheetFunction.Round(totalSum, 2) '지원자의 대학 환산점수
     
    totalSum = totalSum + englishScore + koreanHistoryScore
End Sub


Sub SeoulTech() '서울과학기술대학교

    '======================================
    '2024학년도와 2025학년도가 같다.
    '======================================

    Dim targetRange As Range
    Dim inquiry1ScoreImsi As Double
    Dim inquiry2ScoreImsi As Double
    Dim inquiryAverageImsi As Double
    Dim mathScoreImsi As Double
    
    inquiry1ScoreImsi = 0
    inquiry2ScoreImsi = 0
    inquiryAverageImsi = 0
    mathScoreImsi = 0
    
    Set targetRange = ThisWorkbook.Sheets("수능표준점수최고점").Range("B2:C20")
    
    For i = 0 To weightRatioColCount - 1
    
        If i = 0 And IsNumeric(weightRatio(0, i)) And weightRatio(0, i) > 0 Then
        
           totalSum = totalSum + weightRatio(0, i) * weightRatio(1, i) / subjectSatStdScoreMax(0) * 0.01  '국어
           
           
           Debug.Print "국어 표점최고:" & subjectSatStdScoreMax(0)
           
        ElseIf i = 1 And IsNumeric(weightRatio(0, i)) And weightRatio(0, i) > 0 Then
        
           totalSum = totalSum + weightRatio(0, i) * weightRatio(1, i) / subjectSatStdScoreMax(1) * 0.01 '수학
           
           Debug.Print "수학 표점최고:" & subjectSatStdScoreMax(1)
           
           
           
        ElseIf i = 2 And IsNumeric(weightRatio(0, i)) And weightRatio(0, i) > 0 Then
        
           totalSum = totalSum + weightRatio(0, i) * weightRatio(1, i) / subjectSatStdScoreMax(2) * 0.01  '영어
           
           Debug.Print "영어 표점최고:" & subjectSatStdScoreMax(2)
           
           
        ElseIf i = 4 And IsNumeric(weightRatio(0, i)) And weightRatio(0, i) > 0 Then
        
            If Len(inquiry1SubjectName) > 0 Then
        
                inquiry1ScoreImsi = inquiry1Score / subjectSatStdScoreMax(3)
                
                Debug.Print inquiry1SubjectName & " 표점최고:" & subjectSatStdScoreMax(3)
                
            End If
            
            If Len(inquiry2SubjectName) > 0 Then
                inquiry2ScoreImsi = inquiry2Score / subjectSatStdScoreMax(4)
                
                Debug.Print inquiry2SubjectName & " 표점최고:" & subjectSatStdScoreMax(4)
            End If
            
            weightRatio(1, i) = WorksheetFunction.Average(inquiry1ScoreImsi, inquiry2ScoreImsi)
        
           totalSum = totalSum + weightRatio(0, i) * weightRatio(1, i) * 0.01
        ElseIf i = 15 And IsNumeric(weightRatio(0, 1)) And weightRatio(0, i) > 0 Then
        
        
             inquiryAverageImsi = WorksheetFunction.Average(inquiry1Score / subjectSatStdScoreMax(3), _
                                                                inquiry2Score / subjectSatStdScoreMax(4) _
                                                           )
            mathScoreImsi = mathScore / subjectSatStdScoreMax(1)

        
           totalSum = totalSum + weightRatio(0, i) * WorksheetFunction.Max(mathScoreImsi, inquiryAverageImsi) * 0.01
        End If
    Next i

    totalSum = totalSum * 1000 + koreanHistoryScore
   
End Sub
Sub SNUE() '서울교육대학교

    '======================================
    '2024학년도와 2025학년도가 같다.
    '======================================
    Dim satStdScoreMax As Double
    
'    totalSum = 0
    
    If examYearClassification = "2024학년도 정시" Then
        satStdScoreMax = 160
        
    ElseIf examYearClassification = "2025학년도 정시" Then
    
        satStdScoreMax = 160
    End If
    For i = 0 To weightRatioColCount - 1
    
        If IsNumeric(weightRatio(0, i)) And weightRatio(0, i) > 0 Then
            totalSum = totalSum + weightRatio(1, i) * weightRatio(0, i) / satStdScoreMax
        End If
    Next i
End Sub


Sub SNU() '서울대학교
    '======================================
    '2024학년도와 2025학년도가 같다. 단, 예술대학의 감점이 변화가 있음.
    '======================================

    Dim mathAdd As Double
    mathAdd = 0

    For i = 0 To weightRatioColCount - 1
    
             totalSum = totalSum + weightRatio(0, i) * weightRatio(1, i)
    Next i
    If examYearClassification = "2024학년도 정시" And _
        (targetMajor = "동양화과" Or _
        targetMajor = "서양화과" Or _
        targetMajor = "조소과" Or _
        targetMajor = "공예과" Or _
        targetMajor = "디자인과") Then
    
        Select Case wsStudentData.Range(mathGradeCellAddress).Value
        
            Case 1
            
                mathAdd = 0
            Case 2
                mathAdd = -0.5
            
            Case 3
                mathAdd = -2
            Case 4
                mathAdd = -4
            Case 5
                mathAdd = -6
            Case 6
                mathAdd = -8
            Case 7
                mathAdd = -10
            Case 8
                mathAdd = -12
            Case 9
                mathAdd = -14
        End Select
    
    End If
    
    If examYearClassification = "2024학년도 정시" And _
        targetMajor = "성악과" Then
    
        Select Case wsStudentData.Range(mathGradeCellAddress).Value
        
            Case 1, 2, 3, 4
            
                mathAdd = 0
            Case 5
                mathAdd = -0.4
            Case 6
                mathAdd = -0.8
            Case 7
                mathAdd = -1.2
            Case 8
                mathAdd = -1.6
            Case 9
                mathAdd = -2#
        End Select
    
    End If
    If examYearClassification = "2024학년도 정시" And _
        targetMajor = "작곡과" Then
    
        Select Case wsStudentData.Range(mathGradeCellAddress).Value
        
            Case 1
            
                mathAdd = 0
            Case 2
                mathAdd = -0.5
            
            Case 3
                mathAdd = -1
            Case 4
                mathAdd = -1.5
            Case 5
                mathAdd = -2
            Case 6
                mathAdd = -2.5
            Case 7
                mathAdd = -3
            Case 8
                mathAdd = -3.5
            Case 9
                mathAdd = -4
        End Select
    
    End If
    
    If examYearClassification = "2025학년도 정시" And _
        (targetMajor = "동양화과" Or _
        targetMajor = "서양화과" Or _
        targetMajor = "조소과" Or _
        targetMajor = "공예과") Then
    
        Select Case wsStudentData.Range(mathGradeCellAddress).Value
        
            Case 1
            
                mathAdd = 0
            Case 2
                mathAdd = -0.5
            
            Case 3
                mathAdd = -2
            Case 4
                mathAdd = -4
            Case 5
                mathAdd = -6
            Case 6
                mathAdd = -8
            Case 7
                mathAdd = -10
            Case 8
                mathAdd = -12
            Case 9
                mathAdd = -14
        End Select
    
    End If
    
    If examYearClassification = "2025학년도 정시" And _
        targetMajor = "성악과" Then
    
        Select Case wsStudentData.Range(mathGradeCellAddress).Value
        
            Case 1, 2, 3, 4
            
                mathAdd = 0
            Case 5
                mathAdd = -0.4
            Case 6
                mathAdd = -0.8
            Case 7
                mathAdd = -1.2
            Case 8
                mathAdd = -1.6
            Case 9
                mathAdd = -2
        End Select
    
    End If
    
    If examYearClassification = "2025학년도 정시" And _
        targetMajor = "작곡과" Then
    
        Select Case wsStudentData.Range(mathGradeCellAddress).Value
        
            Case 1
            
                mathAdd = 0
            Case 2
                mathAdd = -0.5
            
            Case 3
                mathAdd = -1
            Case 4
                mathAdd = -1.5
            Case 5
                mathAdd = -2
            Case 6
                mathAdd = -2.5
            Case 7
                mathAdd = -3
            Case 8
                mathAdd = -3.5
            Case 9
                mathAdd = -4
        End Select
    
    End If
    
    totalSum = totalSum + englishScore + koreanHistoryScore + bonusCase9 + bonusCase10 + mathAdd + secondLanguageScore
            


End Sub


Sub UOS() 'University Of Seoul, 서울시립대학교
    '======================================
    '2024학년도와 2025학년도가 같다.
    '======================================

    Dim targetRange As Range
    Dim inquiry1ScoreImsi As Double
    Dim inquiry2ScoreImsi As Double
    inquiry1ScoreImsi = 0
    inquiry2ScoreImsi = 0
    
    Set targetRange = ThisWorkbook.Sheets("수능표준점수최고점").Range("B2:C20")
'    If targetMajor = "영어영문학과" Then
'
'        Debug.Print "이런"
'    End If
    For i = 0 To weightRatioColCount - 1
    
        If i = 0 And IsNumeric(weightRatio(0, i)) And weightRatio(0, i) > 0 Then
        
           totalSum = totalSum + weightRatio(0, i) * weightRatio(1, i) / subjectSatStdScoreMax(0) * 10  '국어
           
        ElseIf i = 1 And IsNumeric(weightRatio(0, i)) And weightRatio(0, i) > 0 Then
        
           totalSum = totalSum + weightRatio(0, i) * weightRatio(1, i) / subjectSatStdScoreMax(1) * 10 '수학
        ElseIf i = 2 And IsNumeric(weightRatio(0, i)) And weightRatio(0, i) > 0 Then
        
           totalSum = totalSum + weightRatio(0, i) * weightRatio(1, i) / subjectSatStdScoreMax(2) * 10  '영어
           
        ElseIf i = 4 And IsNumeric(weightRatio(0, i)) And weightRatio(0, i) > 0 Then
        
            If Len(inquiry1SubjectName) > 0 Then
                inquiry1ScoreImsi = inquiry1Score / subjectSatStdScoreMax(3)
            End If
            If Len(inquiry2SubjectName) > 0 Then
                inquiry2ScoreImsi = inquiry2Score / subjectSatStdScoreMax(4)
            End If
            
            weightRatio(1, i) = inquiry1ScoreImsi + inquiry2ScoreImsi
            
             totalSum = totalSum + weightRatio(0, i) / 2 * weightRatio(1, i) * 10 '탐구는 두 과목 평균이라 /2를 함. 모집요강을 따라 한 것
        
        End If
    Next i
    
'    If targetMajor = "스포츠과학과" Then
'        totalSum = totalSum / 2 + koreanHistoryScore  '만점이 500점
'    Else
'
'        totalSum = totalSum + koreanHistoryScore '만점이 1000점
'    End If
'    If targetMajor = "영어영문학과" Then
'
'        Debug.Print "이런"
'    End If
    
    totalSum = totalSum * wsJeongsi.Cells(jeongsiFoundRow, colJeongsiMaxScore).Value / 1000 + koreanHistoryScore
    
End Sub

Sub SWU() '서울여자대학교
    '======================================
    '2024학년도와 2025학년도가 같다.
    '======================================
    Dim bCount50 As Integer
    
    bCount50 = 0
    
    For i = 0 To weightRatioColCount - 1
    
        If IsNumeric(weightRatio(0, i)) And weightRatio(0, i) > 0 Then
            
            totalSum = totalSum + weightRatio(1, i)
            
            '반영 비율이 50, 50인 과목만 만점이 60점이다. 그래서 50%가 2번 나오는지 확인하기 위한 것
            If weightRatio(0, i) = 50 Then
                bCount50 = bCount50 + 1
            End If
        End If
    Next i

'    If targetMajor = "산업디자인학과" Or _
'       targetMajor = "공예_컬렉터블디자인전공" Or _
'       targetMajor = "스포츠운동과학과" Or _
'       targetMajor = "현대미술전공" Or _
'       targetMajor = "시각디자인전공" Then
'
''       totalSum = weightRatio(1, 11) + weightRatio(1, 12)
'       totalSum = totalSum * wsJeongsi.Cells(jeongsiFoundRow, colJeongsiMaxScore).Value / 200
'
'
'    End If

    If bCount50 = 2 Then
    
       totalSum = totalSum * wsJeongsi.Cells(jeongsiFoundRow, colJeongsiMaxScore).Value / 200
    End If
    
    totalSum = totalSum + koreanHistoryScore

    
End Sub

Sub Sungkonghoe() '성공회대학교
    '======================================
    '2024학년도와 2025학년도가 같다.
    '======================================

    For i = 0 To weightRatioColCount - 1
        totalSum = totalSum + weightRatio(1, i)  '300점 만점으로 계산됨
    Next i

    totalSum = totalSum * wsJeongsi.Cells(jeongsiFoundRow, colJeongsiMaxScore).Value / 300 + koreanHistoryScore

End Sub
Sub SKK()
    '======================================
    '2024학년도와 2025학년도가 같다. 2026학년도는 좀 바뀌네..
    '======================================
    Dim totalSumImsi As Double
'    Dim imsi As Double
    
    
    totalSumImsi = 0
'    imsi = 0
    


    For i = 0 To weightRatioColCount - 1
    
        If i = 3 And IsNumeric(weightRatio(0, i)) And weightRatio(0, i) > 0 Then
        
            totalSum = totalSum + weightRatio(0, i) * weightRatio(1, i) / 20 * 2 ' 탐구 1과목 반영으로 탐구 2과목 반영하는 경우와 같게 하기 위해 2배
            totalSumImsi = totalSumImsi + weightRatioSKK(0, i) * weightRatio(1, i) / 20 * 2
        Else
            totalSum = totalSum + weightRatio(0, i) * weightRatio(1, i) / 20  '예를 들어 국어 반영비율이 40이면 이를 국어표준점수에 2를 곱한다. 40/20
            totalSumImsi = totalSumImsi + weightRatioSKK(0, i) * weightRatio(1, i) / 20
        End If
    Next i
    
        totalSum = (WorksheetFunction.Max(totalSum, totalSumImsi) + koreanHistoryScore) * wsJeongsi.Cells(jeongsiFoundRow, colJeongsiMaxScore).Value / 1000

End Sub

Sub Sungshin() '성신여자대학교

    '======================================
    '2024학년도와 2025학년도가 같다. 2026학년도는 좀 바뀌네..
    '======================================
    
    For i = 0 To weightRatioColCount - 1
    
             totalSum = totalSum + weightRatio(0, i) * weightRatio(1, i) * 0.01
        
    Next i
    
    totalSum = Application.WorksheetFunction.Round(totalSum, 2) '지원자의 대학 환산점수
    totalSum = totalSum * wsJeongsi.Cells(jeongsiFoundRow, colJeongsiMaxScore).Value / 100 + bonusCase4 + koreanHistoryScore


End Sub


Sub Sejong() '세종대학교

    '======================================
    '2024학년도와 2025학년도가 같다. 2026학년도는 좀 바뀌네..
    '======================================
    For i = 0 To weightRatioColCount - 1
    
             totalSum = totalSum + weightRatio(0, i) * weightRatio(1, i) * 0.01 '200점 만점
        
    Next i
    
    totalSum = totalSum * wsJeongsi.Cells(jeongsiFoundRow, colJeongsiMaxScore).Value / 200 + koreanHistoryScore

End Sub



Sub Sookmyung() '숙명여자대학교
    '======================================
    '2024학년도와 2025학년도가 같다. 2026학년도는 좀 바뀌네..
    '======================================
    Dim convertStdScoreMax As Double
    
    
    If examYearClassification = "2024학년도 정시" Then
        convertStdScoreMax = 69.35
    ElseIf examYearClassification = "2025학년도 정시" Then
    
        convertStdScoreMax = 70
    End If
   
    For i = 0 To weightRatioColCount - 1
        If i = 0 And IsNumeric(weightRatio(0, i)) And weightRatio(0, i) > 0 Then
        
           totalSum = totalSum + weightRatio(0, i) * weightRatio(1, i) / subjectSatStdScoreMax(0) * 0.01   '국어
           
        ElseIf i = 1 And IsNumeric(weightRatio(0, i)) And weightRatio(0, i) > 0 Then
        
           totalSum = totalSum + weightRatio(0, i) * weightRatio(1, i) / subjectSatStdScoreMax(1) * 0.01 '수학
        ElseIf i = 2 And IsNumeric(weightRatio(0, i)) And weightRatio(0, i) > 0 Then
        
           totalSum = totalSum + weightRatio(0, i) * weightRatio(1, i) / subjectSatStdScoreMax(2) * 0.01  '영어
        Else
             totalSum = totalSum + weightRatio(0, i) * weightRatio(1, i) / (convertStdScoreMax * 2) * 0.01 '탐구
        
        End If
    Next i
    
    totalSum = totalSum * wsJeongsi.Cells(jeongsiFoundRow, colJeongsiMaxScore).Value + koreanHistoryScore

End Sub


Sub Soongsil() '숭실대학교
    '======================================
    '2024학년도와 2025학년도가 같다. 2026학년도는 좀 바뀌네..
    '======================================
    Dim convertStdScoreMax As Double
    
    
    If examYearClassification = "2024학년도 정시" Then
        convertStdScoreMax = 71.75
    ElseIf examYearClassification = "2025학년도 정시" Then
    
        convertStdScoreMax = 70
    End If
   
    
    For i = 0 To weightRatioColCount - 1
        If i = 0 And IsNumeric(weightRatio(0, i)) And weightRatio(0, i) > 0 Then
        
           totalSum = totalSum + weightRatio(0, i) * weightRatio(1, i) / subjectSatStdScoreMax(0) * 10  '국어
           
        ElseIf i = 1 And IsNumeric(weightRatio(0, i)) And weightRatio(0, i) > 0 Then
        
           totalSum = totalSum + weightRatio(0, i) * weightRatio(1, i) / subjectSatStdScoreMax(1) * 10 '수학
        ElseIf i = 2 And IsNumeric(weightRatio(0, i)) And weightRatio(0, i) > 0 Then
        
           totalSum = totalSum + weightRatio(0, i) * weightRatio(1, i) / subjectSatStdScoreMax(2) * 10  '영어
        Else
             totalSum = totalSum + weightRatio(0, i) * weightRatio(1, i) / (convertStdScoreMax * 2) * 10 '탐구
'             Debug.Print WorksheetFunction.Max(targetRange1)
        
        End If
    Next i
    
    totalSum = totalSum + koreanHistoryScore


End Sub

Sub Ajou() '아주대학교
    '======================================
    '2024학년도와 2025학년도가 같다. 2026학년도는 좀 바뀌네..
    '======================================
    
    For i = 0 To weightRatioColCount - 1
        If i = 0 And IsNumeric(weightRatio(0, i)) And weightRatio(0, i) > 0 Then
        
           totalSum = totalSum + weightRatio(0, i) * weightRatio(1, i) / subjectSatStdScoreMax(0) * 0.01  '국어
           
        ElseIf i = 1 And IsNumeric(weightRatio(0, i)) And weightRatio(0, i) > 0 Then
        
           totalSum = totalSum + weightRatio(0, i) * weightRatio(1, i) / subjectSatStdScoreMax(1) * 0.01 '수학
        ElseIf i = 2 And IsNumeric(weightRatio(0, i)) And weightRatio(0, i) > 0 Then
        
           totalSum = totalSum + weightRatio(0, i) * weightRatio(1, i) / subjectSatStdScoreMax(2) * 0.01  '영어
        
        ElseIf i = 4 And IsNumeric(weightRatio(0, i)) And weightRatio(0, i) > 0 Then
        
             totalSum = totalSum + weightRatio(0, i) * weightRatio(1, i) / (subjectSatStdScoreMax(3) + subjectSatStdScoreMax(4)) * 0.01 '탐구
        
        End If
    Next i
    
    totalSum = totalSum * 1000 + koreanHistoryScore



End Sub
Sub Yonsei() '연세대학교
    '======================================
    '2024학년도와 2025학년도가 같다. 2026학년도는 좀 바뀌네..
    '======================================
    Dim divideSum As Long
    
    divideSum = 0
    
    For i = 0 To weightRatioColCount - 1
            If i = 2 Then
             totalSum = totalSum + weightRatio(1, i) * weightRatio(0, i) / 100
            Else
             totalSum = totalSum + weightRatio(1, i) * weightRatio(0, i) / 200
             
            End If
            
            divideSum = divideSum + weightRatio(0, i)
    Next i
    

    totalSum = totalSum * wsJeongsi.Cells(jeongsiFoundRow, colJeongsiMaxScore).Value / divideSum + koreanHistoryScore
    
End Sub

Sub Yonsei_Mirae() '연세대학교_미래

    '======================================
    '2024학년도와 2025학년도가 같다.
    '======================================
    Dim divideSum As Long
    
    divideSum = 0
    
    For i = 0 To weightRatioColCount - 1
            If i = 2 Then
             totalSum = totalSum + weightRatio(1, i) * weightRatio(0, i) / 100
             
            ElseIf i = 3 Then
             totalSum = totalSum + weightRatio(1, i) * weightRatio(0, i) * 2 / 200
            Else
             totalSum = totalSum + weightRatio(1, i) * weightRatio(0, i) / 200
             
            End If
            
            divideSum = divideSum + weightRatio(0, i)
    Next i
    

    totalSum = totalSum * wsJeongsi.Cells(jeongsiFoundRow, colJeongsiMaxScore).Value / divideSum + koreanHistoryScore



End Sub

Sub Eulji() '을지대학교
    '======================================
    '2024학년도와 2025학년도가 같다.
    '======================================
    
    For i = 0 To weightRatioColCount - 1
             
             totalSum = totalSum + weightRatio(0, i) * weightRatio(1, i) * 0.01
    
    Next i
    
    If wsJeongsi.Cells(jeongsiFoundRow, "D") = "일반전형Ⅰ" Then
        totalSum = 90 + totalSum * 8.1
    
    ElseIf wsJeongsi.Cells(jeongsiFoundRow, "D") = "일반전형Ⅱ" Then
    
        totalSum = totalSum * 10
    End If
    

    totalSum = totalSum + koreanHistoryScore
    
End Sub



Sub Ewha() '이화여자대학교
    '======================================
    '2024학년도와 2025학년도가 같다.
    '======================================
    Dim inquiry1ScoreImsi As Double
    Dim inquiry2ScoreImsi As Double
    
    For i = 0 To weightRatioColCount - 1
        If i = 0 And IsNumeric(weightRatio(0, i)) And weightRatio(0, i) > 0 Then
        
           totalSum = totalSum + weightRatio(0, i) * weightRatio(1, i) / subjectSatStdScoreMax(0) * 0.01  '국어
           
           
        ElseIf i = 1 And IsNumeric(weightRatio(0, i)) And weightRatio(0, i) > 0 Then
        
           totalSum = totalSum + weightRatio(0, i) * weightRatio(1, i) / subjectSatStdScoreMax(1) * 0.01 '수학
           
        ElseIf i = 2 And IsNumeric(weightRatio(0, i)) And weightRatio(0, i) > 0 Then
        
           totalSum = totalSum + weightRatio(0, i) * weightRatio(1, i) / subjectSatStdScoreMax(2) * 0.01  '영어
           
        
        ElseIf i = 4 And IsNumeric(weightRatio(0, i)) And weightRatio(0, i) > 0 Then
        
            If Len(inquiry1SubjectName) > 0 Then
            
                inquiry1ScoreImsi = inquiry1Score / subjectSatStdScoreMax(3)
            End If
            If Len(inquiry2SubjectName) > 0 Then
            
                inquiry2ScoreImsi = inquiry2Score / subjectSatStdScoreMax(4)
            End If
            
            weightRatio(1, i) = inquiry1ScoreImsi + inquiry2ScoreImsi
            
            totalSum = totalSum + weightRatio(0, i) * weightRatio(1, i) / 2 * 0.01 '탐구과목 2과목의 합임. 그래서 과목당 탐구 반영비율의 1/2를 반영 평균으로..
        
        End If
    Next i
    

    totalSum = totalSum * 1000 + koreanHistoryScore
    
End Sub

Sub Incheon() '인천대학교

    '======================================
    '2024학년도와 2025학년도가 같다.
    '======================================

    '1000점 만점으로 계산
    For i = 0 To weightRatioColCount - 1
        If i = 4 And IsNumeric(weightRatio(0, i)) And weightRatio(0, i) > 0 Then
             totalSum = totalSum + weightRatio(0, i) * weightRatio(1, i) / 10 / 2 '탐구
        Else
             totalSum = totalSum + weightRatio(0, i) * weightRatio(1, i) / 10
        End If
     
    Next i
'
'
''    totalSum = totalSum * 5
'
'    If targetMajor = "체육교육과" Then
'        totalSum = totalSum * 0.5
'
'    ElseIf targetMajor = "스포츠과학부" Or _
'            targetMajor = "운동건강학부" Then
'        totalSum = totalSum * 0.4
'    End If
    
    totalSum = totalSum * wsJeongsi.Cells(jeongsiFoundRow, colJeongsiMaxScore).Value / 1000 + koreanHistoryScore
End Sub




Sub Inha() '인하대학교
    '======================================
    '2024학년도와 2025학년도가 같다.
    '단, 한국사 반영 방법이 달라짐.
    '2024학년도는 한국사가 반영비율에 있음.
    '2025학년도는 한국사 점수를 감산
    '======================================
    Dim totalSumImsi As Double
    Dim targetRange As Range
    Dim targetRange1 As Range
    
    Set targetRange = ThisWorkbook.Sheets("수능표준점수최고점").Range("B2:C20")
    Set targetRange1 = ThisWorkbook.Sheets("25정시 변환표준점수").Range("Y2:Y100") ' 이화여자대학교 X열 변환표준 점수의 최고점을 얻기 위함.
    
    totalSumIsi = 0
    
    For i = 0 To weightRatioColCount - 1
        If i = 0 And IsNumeric(weightRatio(0, i)) And weightRatio(0, i) > 0 Then
        
           totalSum = totalSum + weightRatio(0, i) * weightRatio(1, i) / subjectSatStdScoreMax(0)   '국어
           totalSumImsi = totalSumImsi + weightRatioSKK(0, i) * weightRatio(1, i) / subjectSatStdScoreMax(0)   '국어
        
        
        ElseIf i = 1 And IsNumeric(weightRatio(0, i)) And weightRatio(0, i) > 0 Then
        
           totalSum = totalSum + weightRatio(0, i) * weightRatio(1, i) / subjectSatStdScoreMax(1)  '수학
           totalSumImsi = totalSumImsi + weightRatioSKK(0, i) * weightRatio(1, i) / subjectSatStdScoreMax(1)  '수학
        
        ElseIf i = 2 And IsNumeric(weightRatio(0, i)) And weightRatio(0, i) > 0 Then
           totalSum = totalSum + weightRatio(0, i) * weightRatio(1, i) / subjectSatStdScoreMax(2)   '영어
           totalSumImsi = totalSumImsi + weightRatioSKK(0, i) * weightRatio(1, i) / subjectSatStdScoreMax(2)   '영어
        
'        ElseIf i = 3 And IsNumeric(weightRatio(0, i)) And weightRatio(0, i) > 0 Then
'             totalSum = totalSum + weightRatio(0, i) * weightRatio(1, i) / WorksheetFunction.Max(subjectSatStdScoreMax(3), subjectSatStdScoreMax(4))  '탐구
'             totalSumImsi = totalSumImsi + weightRatioSKK(0, i) * weightRatio(1, i) / WorksheetFunction.Max(subjectSatStdScoreMax(3), subjectSatStdScoreMax(4))  '탐구
             
        ElseIf i = 4 And IsNumeric(weightRatio(0, i)) And weightRatio(0, i) > 0 Then
             totalSum = totalSum + weightRatio(0, i) * weightRatio(1, i) / WorksheetFunction.Max(subjectSatStdScoreMax(3), subjectSatStdScoreMax(4))  '탐구
             totalSumImsi = totalSumImsi + weightRatioSKK(0, i) * weightRatio(1, i) / WorksheetFunction.Max(subjectSatStdScoreMax(3), subjectSatStdScoreMax(4))  '탐구
        
        ElseIf i = 5 And IsNumeric(weightRatio(0, i)) And weightRatio(0, i) > 0 Then
             totalSum = totalSum + weightRatio(0, i) * weightRatio(1, i) / 50  '탐구
             totalSumImsi = totalSumImsi + weightRatioSKK(0, i) * weightRatio(1, i) / 50  '탐구
        
        End If
    Next i
        
    If examYearClassification = "2024학년도 정시" Then
    
        totalSum = WorksheetFunction.Max(totalSum, totalSumImsi) * wsJeongsi.Cells(jeongsiFoundRow, colJeongsiMaxScore).Value / 100
        
    ElseIf examYearClassification = "2025학년도 정시" Then
    
        totalSum = WorksheetFunction.Max(totalSum, totalSumImsi) * wsJeongsi.Cells(jeongsiFoundRow, colJeongsiMaxScore).Value / 100 + koreanHistoryScore
    End If
End Sub

Sub Chonnam() '전남대학교
    '======================================
    '2024학년도와 2025학년도가 같다.
    '과목별 표준점수, 변환표준점수 최고점이 달라
    '과목별 표준점수의 변화가 작아도 환산점수는 크게 달라질 수 있음.
    '======================================

    
    For i = 0 To weightRatioColCount - 1
        If i = 0 And IsNumeric(weightRatio(0, i)) And weightRatio(0, i) > 0 Then
        
           totalSum = totalSum + weightRatio(0, i) * weightRatio(1, i) / subjectSatStdScoreMax(0) * 0.01  '국어
           
        ElseIf i = 1 And IsNumeric(weightRatio(0, i)) And weightRatio(0, i) > 0 Then
        
           totalSum = totalSum + weightRatio(0, i) * weightRatio(1, i) / subjectSatStdScoreMax(1) * 0.01  '수학
           
        ElseIf i = 3 And IsNumeric(weightRatio(0, i)) And weightRatio(0, i) > 0 Then
        
            If inquiry1Score >= inquiry2Score Then
            
                totalSum = totalSum + weightRatio(0, i) * weightRatio(1, i) / subjectSatStdScoreMax(3) * 0.01  '탐구
            Else
            
                totalSum = totalSum + weightRatio(0, i) * weightRatio(1, i) / subjectSatStdScoreMax(4) * 0.01  '탐구
             
            End If
        ElseIf i = 4 And IsNumeric(weightRatio(0, i)) And weightRatio(0, i) > 0 Then
        
            If Len(inquiry1SubjectName) > 0 Then
                inquiry1ScoreImsi = inquiry1Score / subjectSatStdScoreMax(3)
            End If
            If Len(inquiry2SubjectName) > 0 Then
                inquiry2ScoreImsi = inquiry2Score / subjectSatStdScoreMax(4)
            End If
            
            weightRatio(1, i) = inquiry1ScoreImsi + inquiry2ScoreImsi
            
            totalSum = totalSum + weightRatio(0, i) * weightRatio(1, i) / 2 * 0.01 '탐구과목 2과목의 합임. 그래서 과목당 탐구 반영비율의 1/2를 반영
        
        End If
    Next i
    
    
    
    totalSum = totalSum * (wsJeongsi.Cells(jeongsiFoundRow, colJeongsiMaxScore).Value - subjectSatStdScoreMax(2)) + englishScore + koreanHistoryScore
End Sub






Sub Jeonbuk() '전북대학교
    '======================================
    '2024학년도와 2025학년도가 같다.
    '======================================
    For i = 0 To weightRatioColCount - 1
    
            totalSum = totalSum + weightRatio(0, i) * weightRatio(1, i)
    
    Next i
    
    totalSum = totalSum + englishScore + koreanHistoryScore
End Sub

Sub CAU() '중앙대학교
    '======================================
    '2024학년도와 2025학년도가 같다.
    '======================================
     
    For i = 0 To weightRatioColCount - 1
    
             totalSum = totalSum + weightRatio(0, i) * weightRatio(1, i) * 0.01 '200점 만점으로 계산됨
        
    Next i
    
    totalSum = totalSum * wsJeongsi.Cells(jeongsiFoundRow, colJeongsiMaxScore).Value / 200 + koreanHistoryScore + englishScore

End Sub


Sub Chungnam() '충남대학교
    '======================================
    '2024학년도와 2025학년도가 같다.
    '======================================
    For i = 0 To weightRatioColCount - 1
    
             totalSum = totalSum + weightRatio(0, i) * weightRatio(1, i)
    Next i
    
    
        totalSum = totalSum / 200 + englishScore + koreanHistoryScore

End Sub


Sub Chungbuk() '충북대학교  예체능계 점수 산출이 문제네...
    
    '======================================
    '2024학년도와 2025학년도가 같다.
    '======================================
    For i = 0 To weightRatioColCount - 1
    
        If i = 0 And IsNumeric(weightRatio(0, i)) And weightRatio(0, i) > 0 Then
        
            totalSum = totalSum + weightRatio(0, i) * 0.01 * weightRatio(1, i) / subjectSatStdScoreMax(0)     '국어
           
        ElseIf i = 1 And IsNumeric(weightRatio(0, i)) And weightRatio(0, i) > 0 Then
        
            totalSum = totalSum + weightRatio(0, i) * 0.01 * weightRatio(1, i) / subjectSatStdScoreMax(1)     '수학
           
        ElseIf i = 2 And IsNumeric(weightRatio(0, i)) And weightRatio(0, i) > 0 Then
           
            totalSum = totalSum + weightRatio(0, i) * 0.01 * weightRatio(1, i) / subjectSatStdScoreMax(2)      '영어
        
        
        ElseIf i = 4 And IsNumeric(weightRatio(0, i)) And weightRatio(0, i) > 0 Then
            
               totalSum = totalSum + weightRatio(0, i) * 0.01 * weightRatio(1, i) / (subjectSatStdScoreMax(3) + subjectSatStdScoreMax(4)) '탐구
            
        End If
    Next i

        totalSum = totalSum * 200 + 800 + koreanHistoryScore '기본점수 800점
End Sub

Sub KNUE() '한국교원대학교
    '======================================
    '2024학년도와 2025학년도가 같다.
    '======================================

    For i = 0 To weightRatioColCount - 1
    
             totalSum = totalSum + weightRatio(0, i) * weightRatio(1, i)
        
    Next i
    
    If wsJeongsi.Cells(jeongsiFoundRow, "D").Value = "예술체육실기" Then

        totalSum = totalSum * 260 / 300
    End If
        
    
End Sub

Sub HUFS() '한국외국어대학교, 한국외국어대학교_글로벌캠퍼스

    Dim convertedStdScoreMax As Double

    '======================================
    '2024학년도와 2025학년도가 같다.
    '======================================
    
    If examYearClassification = "2024학년도 정시" Then
    
        convertedStdScoreMax = 69.35
        
    ElseIf examYearClassification = "2025학년도 정시" Then
    
        convertedStdScoreMax = 70
    End If
    
    For i = 0 To weightRatioColCount - 1
    
        If i = 0 And IsNumeric(weightRatio(0, i)) And weightRatio(0, i) > 0 Then
        
           totalSum = totalSum + weightRatio(0, i) * weightRatio(1, i) / subjectSatStdScoreMax(0) * 0.01   '국어
           
        ElseIf i = 1 And IsNumeric(weightRatio(0, i)) And weightRatio(0, i) > 0 Then
        
           totalSum = totalSum + weightRatio(0, i) * weightRatio(1, i) / subjectSatStdScoreMax(1) * 0.01  '수학
        
        ElseIf i = 4 And IsNumeric(weightRatio(0, i)) And weightRatio(0, i) > 0 Then
        
            '과목의 표준점수의 최고점이 아니라 변화표준점수의 최고점이라 수정
'             totalSum = totalSum + weightRatio(0, i) * weightRatio(1, i) / (WorksheetFunction.Max(subjectSatStdScoreMax(3), subjectSatStdScoreMax(4)) * 2) * 0.01 '탐구
             totalSum = totalSum + weightRatio(0, i) * weightRatio(1, i) / (convertedStdScoreMax * 2) * 0.01 '탐구
        End If
    Next i
    
    totalSum = totalSum * 700 + englishScore + koreanHistoryScore ' 입시요강에 700점을 곱한다고 되어있음.

End Sub

Sub Hanyang() '한양대학교

    '======================================
    '2024학년도와 2025학년도가 같다.
    '국어, 영어 표준 점수를 나누는 분모 값은 해당 학년도 수능의 각 과목 최고 점수이고
    ' 탐구를 나누는 분모의 값은 해당 학년도  변환표준점수 최고점이다
    '======================================
    Dim inquiryMax As Double
    
    If examYearClassification = "2024학년도 정시" Then
    
        inquiryMax = 68.85
    
    ElseIf examYearClassification = "2025학년도 정시" Then
    
        inquiryMax = 70
    End If
    
    
    For i = 0 To weightRatioColCount - 1
        If i = 0 And IsNumeric(weightRatio(0, i)) And weightRatio(0, i) > 0 Then
        
           totalSum = totalSum + weightRatio(0, i) * weightRatio(1, i) / subjectSatStdScoreMax(0) * 10  '국어
           
        ElseIf i = 1 And IsNumeric(weightRatio(0, i)) And weightRatio(0, i) > 0 Then
        
           totalSum = totalSum + weightRatio(0, i) * weightRatio(1, i) / subjectSatStdScoreMax(1) * 10 '수학
        ElseIf i = 2 And IsNumeric(weightRatio(0, i)) And weightRatio(0, i) > 0 Then
        
           totalSum = totalSum + englishScore
        ElseIf i = 3 And IsNumeric(weightRatio(0, i)) And weightRatio(0, i) > 0 Then
        
             totalSum = totalSum + weightRatio(0, i) * weightRatio(1, i) / inquiryMax * 10 '탐구 1과목 반영
             
        ElseIf i = 4 And IsNumeric(weightRatio(0, i)) And weightRatio(0, i) > 0 Then
             totalSum = totalSum + weightRatio(0, i) * weightRatio(1, i) / (inquiryMax * 2) * 10 '탐구 2과목 반영
        
        End If
    Next i
    
    totalSum = totalSum * wsJeongsi.Cells(jeongsiFoundRow, colJeongsiMaxScore).Value / 1000 + koreanHistoryScore

End Sub


Sub HanyangErica() '한양대학교_에리카

    '======================================
    '2024학년도와 2025학년도가 같다.
    '국어, 영어 표준 점수를 나누는 분모 값은 해당 학년도 수능의 각 과목 최고 점수이고
    ' 탐구를 나누는 분모의 값은 해당 학년도  변환표준점수 최고점이다
    '======================================
    Dim inquiryMax As Double
    
    If examYearClassification = "2024학년도 정시" Then
    
        inquiryMax = 68.85
    
    ElseIf examYearClassification = "2025학년도 정시" Then
    
        inquiryMax = 70
    End If
    
    For i = 0 To weightRatioColCount - 1
    
        If i = 0 And IsNumeric(weightRatio(0, i)) And weightRatio(0, i) > 0 Then
        
           totalSum = totalSum + weightRatio(0, i) * weightRatio(1, i) / subjectSatStdScoreMax(0) * 10  '국어
           
        ElseIf i = 1 And IsNumeric(weightRatio(0, i)) And weightRatio(0, i) > 0 Then
        
           totalSum = totalSum + weightRatio(0, i) * weightRatio(1, i) / subjectSatStdScoreMax(1) * 10 '수학
           
        ElseIf i = 2 And IsNumeric(weightRatio(0, i)) And weightRatio(0, i) > 0 Then
        
            totalSum = totalSum + englishScore
            
        ElseIf i = 3 And IsNumeric(weightRatio(0, i)) And weightRatio(0, i) > 0 Then
        
             totalSum = totalSum + weightRatio(0, i) * weightRatio(1, i) / inquiryMax * 10 '탐구 1과목반영
             
        ElseIf i = 4 And IsNumeric(weightRatio(0, i)) And weightRatio(0, i) > 0 Then
        
             totalSum = totalSum + weightRatio(0, i) * weightRatio(1, i) / (inquiryMax * 2) * 10 '탐구 2과목반영
             
        ElseIf i = 15 And IsNumeric(weightRatio(0, i)) And weightRatio(0, i) > 0 Then
        
            totalSum = totalSum + WorksheetFunction.Max(mathScore / subjectSatStdScoreMax(1), _
                                                      WorksheetFunction.Max(inquiry1Score, inquiry2Score) / inquiryMax _
                                                      ) _
                                                      * weightRatio(0, i) * 10
                                                    
        
        End If
        
    Next i
    
    totalSum = totalSum + koreanHistoryScore


End Sub






Sub Kyungpook() '경북대학교

    '====================================================================
    '2024학년도와 2025학년도의 반영비율이 조금 변경되었으나
    '유형 ABC의 구분과 각 구분별 반영법은 동일
    '유형 C만의 독특한 점은 수학의 점수가 400이고 탐구1과목의 점수가 200이다. 이를 기준으로 구분
    '====================================================================
    Dim imsi As Integer
    imsi = 0

    For i = 0 To weightRatioColCount - 1
      
             
        If i = 2 And IsNumeric(weightRatio(0, i)) And weightRatio(0, i) > 0 Then
        
            If weightRatio(0, i) = 200 Then
            
                totalSum = totalSum + weightRatio(1, i) * 2 '영어 유형C로 영어를 2배 반영
            Else
            
                totalSum = totalSum + weightRatio(1, i)  '영어 유형C가 아님.
            End If
        ElseIf i = 3 And IsNumeric(weightRatio(0, i)) And weightRatio(0, i) > 0 Then
        
             totalSum = totalSum + weightRatio(0, i) * weightRatio(1, i) / 200 * 2 ' 유형C로 탐구과목이 1개만 반영되어서 2배
        
        ElseIf IsNumeric(weightRatio(0, i)) And weightRatio(0, i) > 0 Then
        
             totalSum = totalSum + weightRatio(0, i) * weightRatio(1, i) / 200
        
        End If
        
        imsi = imsi + weightRatio(0, i)
        
    Next i
    
    totalSum = totalSum * wsJeongsi.Cells(jeongsiFoundRow, colJeongsiMaxScore).Value / imsi + koreanHistoryScore


End Sub

Sub Gyeongsang() '경상국립대학교
    '========================================
    '2024학년도와 2025학년도의 변화가 없음.
    '========================================
    Dim inquiry1ScoreImsi As Double
    Dim inquiry2ScoreImsi As Double
    Dim targetRange As Range
    
'    Set targetRange = ThisWorkbook.Sheets("수능표준점수최고점").Range("B2:C20")
    inquiry1ScoreImsi = 0
    inquiry2ScoreImsi = 0

    For i = 0 To weightRatioColCount - 1
        If i = 0 And IsNumeric(weightRatio(0, i)) And weightRatio(0, i) > 0 Then
        
           totalSum = totalSum + weightRatio(0, i) * weightRatio(1, i) / subjectSatStdScoreMax(0)
        ElseIf i = 1 And IsNumeric(weightRatio(0, i)) And weightRatio(0, i) > 0 Then
        
           totalSum = totalSum + weightRatio(0, i) * weightRatio(1, i) / subjectSatStdScoreMax(1)  '수학
        ElseIf i = 2 And IsNumeric(weightRatio(0, i)) And weightRatio(0, i) > 0 Then
            If weightRatio(0, i) = 180 Then

                totalSum = totalSum + weightRatio(1, i) * 0.6 '영어
                    
            
            Else
            
                totalSum = totalSum + weightRatio(1, i)  '영어
        
            End If
             
        ElseIf i = 4 And IsNumeric(weightRatio(0, i)) And weightRatio(0, i) > 0 Then
        
            If Len(inquiry1SubjectName) > 0 Then
                inquiry1ScoreImsi = inquiry1Score / subjectSatStdScoreMax(3)
            End If
            If Len(inquiry2SubjectName) > 0 Then
                inquiry2ScoreImsi = inquiry2Score / subjectSatStdScoreMax(4)
            End If
            
            weightRatio(1, i) = WorksheetFunction.Average(inquiry1ScoreImsi, inquiry2ScoreImsi)
            
            totalSum = totalSum + weightRatio(0, i) * weightRatio(1, i)  '탐구과목 2과목의 합임. 그래서 과목당 탐구 반영비율의 1/2를 반영
        
        End If
    Next i

    totalSum = totalSum + weightRatio(1, 1) * (bonusCase2 + bonusCase3) + weightRatio(1, 4) * (bonusCase7 + bonusCase9 + bonusCase10)
End Sub


Sub Kangwon_Chuncheon() '강원대학교_춘천

    If examYearClassification = "2024학년도 정시" Or _
        examYearClassification = "2025학년도 정시" Then
        
        For i = 0 To weightRatioColCount - 1
                 totalSum = totalSum + weightRatio(0, i) * weightRatio(1, i) / 100
            
        Next i
        
        If targetMajor = "체육교육과" Then
            totalSum = totalSum * 3 + koreanHistoryScore
        Else
        
            totalSum = totalSum * 5 + koreanHistoryScore
        End If
    End If
    


End Sub




Sub Pusan()
    '==========================================
    '2024학년도와 2025학년도의 반영 방법 동일
    '==========================================

    For i = 0 To weightRatioColCount - 1
        If i <> 2 Then
             totalSum = totalSum + weightRatio(0, i) * weightRatio(1, i) / 200 '200점 만점으로 계산
        End If
    Next i

    totalSum = totalSum + englishScore + koreanHistoryScore
End Sub

Sub MMU() '국립목포해양대 Mokpo National Marine University

    For i = 0 To weightRatioColCount - 1
    
        totalSum = totalSum + weightRatio(0, i) * weightRatio(1, i) * 0.01
        
    Next i

    totalSum = totalSum * 10

End Sub



Sub UniversityPossibilityProcess()

    Dim scoreGap As Double

    
    universityPossibilityExplain = ""
    '==========================================================================
    '목표 대학, 학과의 만점, 50등, 70등, 내 성적을 변수에 할당
    '==========================================================================
    universityMaxScore = 0
    universityScore50Pct = 0
    universityScore70Pct = 0
    universityMyScore = 0
    universityPossibility = 0
    university5070Difference = 0
    
    If IsNumeric(wsJeongsi.Cells(jeongsiFoundRow, colJeongsiMaxScore).Value) And wsJeongsi.Cells(jeongsiFoundRow, colJeongsiMaxScore).Value > 0 Then
        universityMaxScore = wsJeongsi.Cells(jeongsiFoundRow, colJeongsiMaxScore).Value   '대학발표 환산점 만점
    End If
    
    If IsNumeric(wsJeongsi.Cells(jeongsiFoundRow, colJeongsiScore50Pct).Value) And wsJeongsi.Cells(jeongsiFoundRow, colJeongsiScore50Pct).Value > 0 Then
        universityScore50Pct = wsJeongsi.Cells(jeongsiFoundRow, colJeongsiScore50Pct).Value   '대학발표 합격생 50%(100명 중 50등) 점수
    End If
    
    If IsNumeric(wsJeongsi.Cells(jeongsiFoundRow, colJeongsiScore70Pct).Value) And wsJeongsi.Cells(jeongsiFoundRow, colJeongsiScore70Pct).Value > 0 Then
        universityScore70Pct = wsJeongsi.Cells(jeongsiFoundRow, colJeongsiScore70Pct).Value   '대학발표 합격생 70%(100명 중 70등) 점수
    End If
    
    universityMyScore = Application.WorksheetFunction.Round(totalSum, 5) '지원자의 대학 환산점수
    
    
    ' 50%, 70% 점수가 모두 존재하는 경우
    '   50%, 70% 점수가 같은 경우에는
    '       내 점수가 50% 점수보다 높을 때, 합격확률을 70% 이상으로
    '       50%점수-3.3<= 내 점수 < 50% 점수, 50% 이상  '3.3은 50%점수와 70% 점수 차이의 평균임.
    '       내 점수 < 50% 점수, 50% 미만
    '   내 점수가 50% 점수 이상일 때,
    '       내점수>= 50% 점수 + (50%d점수- 70% 점수)이면
    '           80% 이상
    '       50% 점수 <= 내점수 < 50% 점수 + (50%d점수- 70% 점수)
    '           70% 이상
    '   내 점수가 70% 점수 이하일 때,
    '       내점수 <=70% 점수 - (50%d점수- 70% 점수)
    '           30% 미만
    '       내점수 <=70% 점수
    '           50% 미만
    '   70% 점수 <내 점수 < 50% 점수
    '       선형 비례로 계산한다.
    '50% 점수 만 존재하는 경우
    '   내점수 > 50% 점수+3.3
    '       80% 이상
    '   내점수 >=  50%점수
    '       70% 이상
    '   내 점수 >= 50% 점수-3.3
    '       50% 이상
    '   내 점수 >= 50% 점수 -3.3*2
    '       30% 이상
    '   내 점수 < 50% 점수-3.3*2
    '       30% 미만
    '70% 점수 만 존재하는 경우
    '
    universityPossibilityExplain = ""
    
    scoreGap = 1.65
    
    If IsNumeric(universityScore50Pct) And universityScore50Pct > 0 And _
       IsNumeric(universityScore70Pct) And universityScore70Pct > 0 Then
       
        university5070Difference = universityScore50Pct - universityScore70Pct
        '50%, 70% 점수가 같은 경우
        If universityScore50Pct = universityScore70Pct Then
       
            If universityMyScore >= universityScore50Pct + 2 * scoreGap Then
            
                universityPossibilityExplain = "80% 이상"
                
            ElseIf universityMyScore >= universityScore50Pct + scoreGap Then
            
                universityPossibilityExplain = "70% 이상"
                
            ElseIf universityMyScore >= (universityScore50Pct) Then
            
                universityPossibilityExplain = "60% 이상"
                
            ElseIf universityMyScore >= (universityScore50Pct - scoreGap) Then
            
                universityPossibilityExplain = "50% 이상"
                
            ElseIf universityMyScore >= (universityScore50Pct - scoreGap * 2) Then
            
                universityPossibilityExplain = "40% 이상"
                
            ElseIf universityMyScore >= (universityScore50Pct - scoreGap * 3) Then
            
                universityPossibilityExplain = "30% 이상"
            Else
               universityPossibilityExplain = "30% 미만"
             
            End If
       
        ElseIf universityMyScore >= universityScore50Pct Then
       
            If universityMyScore >= (universityScore50Pct + university5070Difference) Then
                
                universityPossibilityExplain = "90% 이상"
                
            ElseIf universityMyScore >= (universityScore50Pct + university5070Difference / 2) Then
                
                universityPossibilityExplain = "80% 이상"
            Else
                universityPossibilityExplain = "70% 이상"
            End If
        
        ElseIf universityMyScore >= universityScore70Pct Then
       
            universityPossibility = 60 + 20 * (universityMyScore - universityScore70Pct) / (universityScore50Pct - universityScore70Pct)
'            t = WorksheetFunction.Round(UniversityPossibility, -1)
            
            universityPossibilityExplain = WorksheetFunction.Round(universityPossibility, -1) & "% 이상"
      
        
        Else
            
            If universityMyScore >= (universityScore70Pct - university5070Difference / 2) Then
            
                universityPossibilityExplain = "50% 이상"
                
            ElseIf universityMyScore >= (universityScore70Pct - university5070Difference) Then
            
                universityPossibilityExplain = "40% 이상"
                
            ElseIf universityMyScore >= (universityScore70Pct - university5070Difference * 3 / 2) Then
            
                universityPossibilityExplain = "30% 이상"
            Else
                universityPossibilityExplain = "30% 미만"
            End If
            
       
        End If
    
    ElseIf IsNumeric(universityScore50Pct) And universityScore50Pct > 0 Then
        
        If universityMyScore >= (universityScore50Pct + scoreGap) Then
        
            universityPossibilityExplain = "80% 이상"
        
        ElseIf universityMyScore >= universityScore50Pct Then

            universityPossibilityExplain = "70% 이상"

        ElseIf universityMyScore >= (universityScore50Pct - scoreGap) Then ''50%, 70% 합격생의 평균점수의 차이를 2라고 가정했다
        
            universityPossibilityExplain = "60% 이상"
        ElseIf universityMyScore >= (universityScore50Pct - 2 * scoreGap) Then ''50%, 70% 합격생의 평균점수의 차이를 2라고 가정했다
            universityPossibilityExplain = "50% 이상"
            
        ElseIf universityMyScore >= (universityScore50Pct - 3 * scoreGap) Then ''50%, 70% 합격생의 평균점수의 차이를 2라고 가정했다
            universityPossibilityExplain = "40% 이상"
            
        ElseIf universityMyScore >= (universityScore50Pct - 4 * scoreGap) Then ''50%, 70% 합격생의 평균점수의 차이를 2라고 가정했다
            universityPossibilityExplain = "30% 이상"
        Else
        
            universityPossibilityExplain = "30% 미만"
            
        End If
    
    
    ElseIf IsNumeric(universityScore70Pct) And universityScore70Pct > 0 Then
    
        If universityMyScore >= (universityScore70Pct + 3 * scoreGap) Then
        
            universityPossibilityExplain = "90% 이상"

        ElseIf universityMyScore >= (universityScore70Pct + 2 * scoreGap) Then
        
            universityPossibilityExplain = "80% 이상"
        ElseIf universityMyScore >= (universityScore70Pct + scoreGap) Then
        
            universityPossibilityExplain = "70% 이상"
        ElseIf universityMyScore >= universityScore70Pct Then '
        
            universityPossibilityExplain = "60% 이상"
        
        ElseIf universityMyScore >= (universityScore70Pct - scoreGap) Then
        
            universityPossibilityExplain = "50% 이상"
        ElseIf universityMyScore >= (universityScore70Pct - 2 * scoreGap) Then
        
            universityPossibilityExplain = "40% 이상"
            
        ElseIf universityMyScore >= (universityScore70Pct - 3 * scoreGap) Then
        
            universityPossibilityExplain = "30% 이상"
        Else
        
            universityPossibilityExplain = "30% 미만"
            
        End If
          
    End If
    
End Sub




Sub DisplayConvertScoreResult()


    If worksheetName = "대학 학과별 상세검색" Then
    
        wsStudentData.Cells(13, "F") = koreanScoreDisplay
        wsStudentData.Cells(13, "G") = koreanScoreProcessType
        wsStudentData.Cells(14, "F") = mathScoreDisplay
        wsStudentData.Cells(14, "G") = mathScoreProcessType
        wsStudentData.Cells(15, "F") = englishScoreDisplay
        wsStudentData.Cells(15, "G") = englishScoreProcessType
        wsStudentData.Cells(16, "F") = koreanHistoryScoreDisplay
        wsStudentData.Cells(16, "G") = koreanHistoryScoreProcessType
        
        wsStudentData.Cells(17, "F") = inquiry1ScoreDisplay
        wsStudentData.Cells(17, "G") = inquiry1ScoreProcessType
        wsStudentData.Cells(18, "F") = inquiry2ScoreDisplay
        wsStudentData.Cells(18, "G") = inquiry2ScoreProcessType
        
        If targetUniversity = "서울대학교" And IsNumeric(secondLanguageScore) And secondLanguageScore > 0 Then
            wsStudentData.Cells(19, "F") = secondLanguageScoreDisplay
            wsStudentData.Cells(19, "G") = secondLanguageScoreProcessType
        End If
    
        '==========================================================================
        '목표 대학, 학과의 만점, 50등, 70등, 내 성적을 적는다.
        '==========================================================================
        wsStudentData.Cells(8, "D") = wsJeongsi.Cells(jeongsiFoundRow, colJeongsiMaxScore).Value   '대학발표 환산점 만점
        wsStudentData.Cells(8, "E") = wsJeongsi.Cells(jeongsiFoundRow, colJeongsiScore50Pct).Value   '대학발표 합격생 50%(100명 중 50등) 점수
        wsStudentData.Cells(8, "F") = wsJeongsi.Cells(jeongsiFoundRow, colJeongsiScore70Pct).Value   '대학발표 합격생 70%(100명 중 70등) 점수
'        wsStudentData.Cells(8, "G") = Application.WorksheetFunction.Round(totalSum, 5) '지원자의 대학 환산점수
        wsStudentData.Cells(8, "G") = Application.WorksheetFunction.Round(totalSum, 3) '지원자의 대학 환산점수
        wsStudentData.Cells(8, "H") = universityPossibilityExplain
    
    
    ElseIf worksheetName = "대학별검색" Then
    
        wsStudentData.Cells(resultDisplayRowNo, "A") = targetUniversity  '대학발표 환산점 만점
        wsStudentData.Cells(resultDisplayRowNo, "B") = wsJeongsi.Cells(jeongsiFoundRow, "D").Value
        wsStudentData.Cells(resultDisplayRowNo, "C") = targetMajor  '대학발표 환산점 만점
        wsStudentData.Cells(resultDisplayRowNo, "D") = wsJeongsi.Cells(jeongsiFoundRow, "F").Value
        
        wsStudentData.Cells(resultDisplayRowNo, "E") = wsJeongsi.Cells(jeongsiFoundRow, "G").Value
        wsStudentData.Cells(resultDisplayRowNo, "F") = wsJeongsi.Cells(jeongsiFoundRow, "H").Value
        wsStudentData.Cells(resultDisplayRowNo, "G") = wsJeongsi.Cells(jeongsiFoundRow, colJeongsiMaxScore).Value  '대학발표 환산점 만점
        wsStudentData.Cells(resultDisplayRowNo, "H") = wsJeongsi.Cells(jeongsiFoundRow, colJeongsiScore50Pct).Value   '대학발표 합격생 50%(100명 중 50등) 점수
        wsStudentData.Cells(resultDisplayRowNo, "I") = wsJeongsi.Cells(jeongsiFoundRow, colJeongsiScore70Pct).Value   '대학발표 합격생 70%(100명 중 70등) 점수
'        wsStudentData.Cells(8, "G") = Application.WorksheetFunction.Round(totalSum, 5) '지원자의 대학 환산점수
        wsStudentData.Cells(resultDisplayRowNo, "J") = Application.WorksheetFunction.Round(totalSum, 3) '지원자의 대학 환산점수
        wsStudentData.Cells(resultDisplayRowNo, "K") = universityPossibilityExplain
        wsStudentData.Cells(resultDisplayRowNo, "L") = wsJeongsi.Cells(jeongsiFoundRow, "DB").Value

    
    End If

    
End Sub

Sub EraseBorderLines()

    Dim targetRange As Range
    Dim bordersReferencePoint As String
    
    
    ' A21 셀을 기준으로 데이터가 있는 연속된 영역(CurrentRegion)을 찾아서 targetRange 변수에 할당해.
    ' CurrentRegion은 A21 셀을 포함하여 빈 행이나 열로 구분되지 않는 모든 셀 범위를 의미해.
    bordersReferencePoint = ""
    
    If worksheetName = "대학별검색" Then
            bordersReferencePoint = "A16"

    Else
         bordersReferencePoint = "A21"
    
    End If
    
    Set targetRange = wsStudentData.Range(bordersReferencePoint).CurrentRegion
    
    ' 이제 찾은 targetRange에 테두리를 그려줄 거야.
    ' 이전 테두리가 있을 수도 있으니 먼저 모든 테두리를 지우고 시작하는 게 좋아.
    targetRange.Select
    targetRange.Borders.LineStyle = xlNone


End Sub
Sub ShowBorderLines()

    Dim targetRange As Range
    Dim bordersReferencePoint As String
    
    ' A21 셀을 기준으로 데이터가 있는 연속된 영역(CurrentRegion)을 찾아서 targetRange 변수에 할당해.
    ' CurrentRegion은 A21 셀을 포함하여 빈 행이나 열로 구분되지 않는 모든 셀 범위를 의미해.
    
    bordersReferencePoint = ""
    
    If worksheetName = "대학별검색" Then
            bordersReferencePoint = "A16"

    Else
         bordersReferencePoint = "A21"
    
    End If
    
    
    
    Set targetRange = wsStudentData.Range(bordersReferencePoint).CurrentRegion
    
    ' 이제 찾은 targetRange에 테두리를 그려줄 거야.
    ' 이전 테두리가 있을 수도 있으니 먼저 모든 테두리를 지우고 시작하는 게 좋아.
    targetRange.Borders.LineStyle = xlNone
    
    ' 그리고 원하는 스타일의 테두리를 적용해줘.
    With targetRange.Borders
        .LineStyle = xlContinuous ' 실선 (일반적으로 사용되는 선 스타일)
        .Weight = xlThin          ' 선 두께를 가늘게 (xlThin, xlMedium, xlThick 중에서 선택 가능)
        .Color = vbBlack          ' 선 색깔을 검정색으로 (vbBlack, vbRed, vbBlue 등 다양한 색상 가능)
    End With
    
    '        MsgBox "A21부터 시작하는 연속된 영역에 윤곽선을 그렸어!", vbInformation

End Sub

Sub DrawStudentScore()

    Dim wsSource As Worksheet
    Dim wsResult As Worksheet
    
    Set wsSource = ThisWorkbook.Sheets("대학별검색")
    Set wsResult = ThisWorkbook.Sheets("대학 학과별 상세검색")

    wsResult.Range("C2").Value = wsSource.Range("C3")
    wsResult.Range("C3").Value = wsSource.Range("D3")
    wsResult.Range("C4").Value = wsSource.Range("E3")
'    wsResult.Range("C7").Value = wsSource.Range()
'    wsResult.Range("C8").Value = wsSource.Range()
    
    
    wsResult.Range("B13").Value = wsSource.Range("B5")
    wsResult.Range("C13").Value = wsSource.Range("C5")
    wsResult.Range("D13").Value = wsSource.Range("D5")
    wsResult.Range("E13").Value = wsSource.Range("E5")
    
    

    wsResult.Range("B14").Value = wsSource.Range("B6")
    wsResult.Range("C14").Value = wsSource.Range("C6")
    wsResult.Range("D14").Value = wsSource.Range("D6")
    wsResult.Range("E14").Value = wsSource.Range("E6")
    
    
    wsResult.Range("E15").Value = wsSource.Range("E7")
    wsResult.Range("E16").Value = wsSource.Range("E8")
    
    
    wsResult.Range("B17").Value = wsSource.Range("B9")
    wsResult.Range("C17").Value = wsSource.Range("C9")
    wsResult.Range("D17").Value = wsSource.Range("D9")
    wsResult.Range("E17").Value = wsSource.Range("E9")
    
    
    
    wsResult.Range("B18").Value = wsSource.Range("B10")
    wsResult.Range("C18").Value = wsSource.Range("C10")
    wsResult.Range("D18").Value = wsSource.Range("D10")
    wsResult.Range("E18").Value = wsSource.Range("E10")
    
    wsResult.Range("B19").Value = wsSource.Range("B11")
    wsResult.Range("E19").Value = wsSource.Range("E11")
    
    
    
    
End Sub
