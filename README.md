# TiaApiServer

Siemens TIA Portal Openness API를 HTTP REST로 래핑한 서드파티 브릿지 서버입니다.

TIA Portal의 C# 전용 API를 HTTP REST API로 개방하여, 언어나 도구에 관계없이 PLC 프로젝트를 자동화할 수 있습니다.

```
Siemens TIA Portal (공식)
    ↕ Openness API (Siemens 제공)
TiaApiServer.exe (서드파티 브릿지)
    ↕ REST API (HTTP)
브라우저 / Python / JavaScript / curl / AI 도구
```

## 워크플로우

### 1. 사전 요구사항

| 항목 | 설명 |
|---|---|
| **TIA Portal V20** | 로컬 PC에 설치 필수 |
| **.NET Framework 4.8** | 런타임 필수 |
| **Siemens TIA Openness** | Windows 그룹 'Siemens TIA Openness'에 사용자 등록 |
| **TiaApiServer.exe** | 이 프로젝트를 빌드하거나 exe 파일을 받기 |

### 2. 실행

```
TiaApiServer.exe 더블클릭 (또는 커맨드라인에서 실행)
→ localhost:8099 에서 HTTP 서버 시작
→ Swagger UI + REST API가 동시에 제공됩니다
```

포트를 변경하려면 인자로 전달합니다: `TiaApiServer.exe 9000`

### 3. 연결

```
1. TIA Portal 실행 → 프로젝트 열기
2. 브라우저에서 localhost:8099 접속
3. /api/connect 호출 → TIA Portal 프로세스에 Attach
```

### 4. API 사용

| 기능 | 설명 |
|---|---|
| **조회** | 프로젝트 → 디바이스 → PLC 블록/태그/타입 탐색 |
| **내보내기 (Export)** | 블록을 SCL/DB/STL 소스 파일 또는 XML ZIP으로 다운로드 |
| **가져오기 (Import)** | XML 파일 또는 ZIP을 PLC에 Import (타입→태그→블록 순서 자동 처리) |
| **드라이브** | SINAMICS S210 등 드라이브 파라미터 조회, 모터 설정 확인 |
| **HMI** | 태그 테이블, 텍스트 리스트 조회 |
| **하드웨어 카탈로그** | 모듈 검색, 호환성 확인 |
| **프로젝트 생성** | 새 TIA Portal 인스턴스, 프로젝트, 디바이스 생성 |

### 5. 연동 활용

REST API이므로 아무 도구에서 호출할 수 있습니다:

- **Python 스크립트**로 블록 자동 생성
- **CI/CD 파이프라인**에서 PLC 프로그램 배포
- **웹 대시보드**에서 프로젝트 상태 모니터링
- **AI 도구 (Claude 등)**와 연결하여 PLC 코드 자동화

## API 문서

서버 실행 후 브라우저에서 `http://localhost:8099` 접속하면 Swagger UI로 전체 API를 확인하고 테스트할 수 있습니다.

## 빌드

Visual Studio 2022 또는 MSBuild로 빌드합니다:

```bash
MSBuild TiaApiServer/TiaApiServer.csproj /p:Configuration=Release
```

빌드 결과: `TiaApiServer/bin/Release/TiaApiServer.exe`

**exe 파일 하나만 배포하면 됩니다.** Swagger UI, OpenAPI 스펙이 모두 exe에 내장되어 있습니다.

## 주의사항

- 이 API는 **로컬 PC에서만 동작**합니다. TIA Portal Openness API 자체가 로컬 프로세스 간 통신만 지원합니다.
- 원격 서버에서는 사용할 수 없습니다.
- `Siemens.Engineering.dll`은 TIA Portal 설치 시 자동 감지됩니다. 별도 DLL 배포가 필요 없습니다.

## 라이선스

[MIT License](LICENSE)
