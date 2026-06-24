# WorldShot Log

<p align="right">
  <a href="./README.md"><kbd>日本語</kbd></a>
  <a href="./README.en.md"><kbd>English</kbd></a>
  <a href="./README.ko.md"><kbd>한국어</kbd></a>
</p>

![WorldShot Log](./img/banner.png)

WorldShot Log는 VRChat 사진을 정리하고, 다시 보고, 게시하기 좋게 다듬기 위한 Windows용 데스크톱 앱입니다.

날짜와 월드별 사진 관리에 더해 이미지 보정, 자르기, AI 피사체 선택, 배경 투명 저장, 가림 처리, 텍스트 추가, 이미지 오버레이 등의 편집 기능을 제공합니다. 원본 이미지는 덮어쓰지 않으며, 편집한 이미지는 별도 파일로 저장됩니다.

> WorldShot Log는 VRChat Inc.의 공식 앱이 아닙니다.

## 주요 기능

### 사진 정리

- VRChat 사진 드래그 앤 드롭 가져오기
- 등록 폴더 재스캔
- 연도, 월, 날짜별 사진 목록
- 월드별 사진 목록
- 즐겨찾기, 방향, 라벨, World 이름으로 필터링
- 다중 선택, Shift 범위 선택, 드래그 선택
- 목록 화면에서 키보드 이동

### 사진 상세 확인

- 사진을 카드 형식으로 보기
- 사진 상세 모달에서 크게 보기
- 이전/다음 이미지로 이동
- 원본 이미지 열기
- 저장 폴더 열기
- VRChat 페이지 열기
- 즐겨찾기, 라벨, 메모 관리
- 월드 이름과 World URL 수동 편집

### 월드 정보 보조 관리

- 사진에서 World ID를 다루기 쉽게 정리
- 가져오기 후 World 정보 자동 취득
- 같은 World ID의 사진에서는 이미 취득한 정보를 재사용
- 월드 이름, 설명, 태그, World URL 표시

### 앱 내 이미지 편집

- 사진 상세 화면에서 이미지 편집 화면 열기
- 원본을 덮어쓰지 않고 편집 이미지를 별도 저장
- 밝기, 노출, 대비, 하이라이트, 섀도우, 화이트, 블랙 조정
- 색온도, 색조, 채도, 자연스러운 채도 조정
- 선명도, 텍스처, 페이드, 그레인, 비네팅 조정
- 스마트 자동 보정, 학습 보정, 보정 강도 조정
- RGB / HSV 톤 커브와 히스토그램 표시
- 프리셋 적용 및 사용자 프리셋 저장
- 투명 이미지 등을 겹칠 수 있는 이미지 오버레이
- 오버레이 소재 관리, 삭제, 여러 레이어 배치
- 투명도, 합성 방식, 적용 범위, 앞뒤 순서, 크기 조정
- 룰러와 그리드에 가볍게 스냅

### 자르기와 구도 조정

- 원본, 1:1, 16:9, 9:16, 5:4, 4:5, 3:2, 2:3 비율 지원
- VRC 갤러리, 이모지/스티커, 아바타 썸네일용 트리밍 프리셋
- 부족한 영역을 투명하게 남기는 정사각형 캔버스
- 해상도가 지정된 프리셋은 저장 시 지정 크기로 리사이즈
- 줌, 좌우 위치, 상하 위치 조정
- 90도 회전, 자유 회전, 좌우 반전, 상하 반전
- 룰러와 3분할 그리드 표시

### AI 피사체 선택

- 경량 AI로 로컬 피사체 마스크 생성
- 표준 AI withoutBG Snap, 고정밀 AI withoutBG Focus OSS는 필요할 때만 최초 다운로드
- 생성한 피사체 마스크를 보정, 흐림, 텍스트, 이미지 오버레이의 적용 범위로 사용
- 피사체 마스크를 사용한 배경 투명 PNG 저장
- AI 처리는 PC 안에서 실행되며, AI 처리를 위해 이미지가 외부 서버로 전송되지 않습니다
- 사용 모델 및 라이선스 확인: [THIRD_PARTY_NOTICES.md](./THIRD_PARTY_NOTICES.md), [AI Model License Review](./docs/ai-model-license-review.md)

### 흐림 및 가림 처리

- 전체 흐림
- 방사형 흐림
- 흐림, 모자이크, 채우기를 사용한 가림 처리
- 사각형, 원, 자유 그리기 범위 선택
- 적용 전 이동, 크기 조절, 회전
- 사용자 이름, 채팅창, UI, 개인 정보 등을 숨기는 용도에 대응

### 텍스트 추가

- 여러 텍스트 추가
- 글자 크기, 폰트, 굵기, 색상 변경
- 테두리, 그림자, 발광 등 장식
- 글자 간격 조정
- 텍스트 이동 및 회전
- 일본어용 폰트와 장식 폰트
- 폰트 선택 시 각 폰트의 실제 모양을 미리보기
- 최근 사용한 폰트 표시

### 데이터 관리

- 썸네일 재생성
- 누락 파일, 누락 썸네일, World 정보 미취득 확인
- 백업 생성
- 백업에서 복원
- CSV / JSON 내보내기
- 앱 내 제거 안내

## 사용 예

- VRChat 사진 정리
- 촬영한 월드 돌아보기
- X 게시용 이미지 조정
- Booth 게재용 이미지 준비
- 썸네일용 자르기
- 채팅창이나 사용자 이름을 숨긴 게시 이미지 만들기
- VRC 갤러리, 이모지, 스티커용 정사각형 이미지 만들기

## 다운로드

최신 버전은 GitHub Releases에서 다운로드할 수 있습니다.

- [Releases](https://github.com/noma-nomoa/vrchat-world-photo-manager/releases)
- [v2.4.2 릴리스 노트](./release-notes/v2.4.2.md)
- [v2.4.1 릴리스 노트](./release-notes/v2.4.1.md)
- [v2.4.0 릴리스 노트](./release-notes/v2.4.0.md)
- [v2.3.0 릴리스 노트](./release-notes/v2.3.0.md)
- [v2.2.1 릴리스 노트](./release-notes/v2.2.1.md)
- [v2.2.0 릴리스 노트](./release-notes/v2.2.0.md)
- [v2.1.0 릴리스 노트](./release-notes/v2.1.0.md)
- [v2.0.0 릴리스 노트](./release-notes/v2.0.0.md)

Windows용 배포 파일은 `WorldShotLogSetup.exe`입니다.

## 설치

1. GitHub Releases에서 최신 `WorldShotLogSetup.exe`를 다운로드합니다.
2. 다운로드한 `WorldShotLogSetup.exe`를 실행합니다.
3. 설치가 완료되면 WorldShot Log가 실행됩니다.

Windows 보안 경고가 표시될 수 있습니다. 내용을 확인한 뒤 실행해 주세요.

## 업데이트

WorldShot Log는 GitHub Releases를 사용한 업데이트 확인을 지원합니다.

- 배포판 앱에서는 새 버전이 공개된 경우 앱 안에서 알림이 표시됩니다.
- 업데이트 후 첫 실행 시 주요 변경 내용을 앱 안에서 확인할 수 있습니다.
- 수동으로 업데이트하려면 GitHub Releases에서 최신 `WorldShotLogSetup.exe`를 다운로드해 실행하세요.
- 개발 실행(`npm start`)에서는 앱 내 업데이트가 동작하지 않습니다.

## 개인정보와 데이터 처리

- 가져온 사진, 메모, 라벨, 설정, 썸네일은 사용자의 PC 안에 저장됩니다.
- 원본 이미지는 덮어쓰지 않습니다. 이미지 편집으로 저장한 파일은 별도 이름으로 저장됩니다.
- World 정보 취득이나 업데이트 확인을 위해 VRChat, GitHub 등의 외부 서비스에 접근할 수 있습니다.
- AI 피사체 선택은 PC 안에서 실행됩니다. 모델을 추가하는 경우에만 모델 파일 다운로드 대상에 접근합니다.
- 추가 다운로드한 AI 모델과 이미지 오버레이 소재는 앱 관리 폴더 안에 저장되며, 앱 안에서 삭제할 수 있습니다.
- 앱에서 만든 백업이나 CSV / JSON 내보내기는 사용자가 선택한 저장 위치에 출력됩니다.
- 제거 시 데이터도 삭제하는 작업을 선택하면 저장된 앱 데이터가 삭제됩니다.

## 데이터 저장 위치

앱 데이터는 주로 다음 위치에 저장됩니다.

- DB / 설정: `C:\Users\<UserName>\AppData\Roaming\WorldShot Log\data\`
- AI 모델: `C:\Users\<UserName>\AppData\Roaming\WorldShot Log\models\`
- 이미지 오버레이 소재: `C:\Users\<UserName>\AppData\Roaming\WorldShot Log\photo-editor-overlays\`
- 썸네일: `C:\Users\<UserName>\WorldShot Log\thumbnails`
- 실행 중 캐시: `C:\Users\<UserName>\AppData\Local\WorldShot Log\SessionData\`

## 이 앱에 대해

WorldShot Log는 개인 제작 앱이며, 설계, 구현, 문서 정리 일부에 AI 지원을 활용하고 있습니다. 최종적인 사양 판단, 동작 확인, 릴리스 판단은 제작자가 진행합니다.

## 관련 문서

- [RELEASE.md](./RELEASE.md): 릴리스 절차
- [AI_MAINTENANCE_GUIDE.md](./AI_MAINTENANCE_GUIDE.md): 유지보수 및 수정 가이드
- [release-notes/](./release-notes): 버전별 변경 내용

## 제한 사항

- private / non-public world의 자동 메타데이터 취득은 지원하지 않습니다.
- 설치 위치를 사용자가 선택하는 기능은 지원하지 않습니다.
- 현재 Windows 이외의 배포는 상정하지 않습니다.

## 라이선스

WorldShot Log는 MIT License로 공개되어 있습니다. 자세한 내용은 [LICENSE](./LICENSE)를 확인하세요.
