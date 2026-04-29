# MVR 룰 스키마

각 룰은 다음 필드를 가진다.

## 공통 필드

| 필드 | 타입 | 필수 | 설명 |
|---|---|---|---|
| `id` | string | ✓ | 룰 고유 ID. 형식: `RULE-MVR-NNN` (3자리 숫자) |
| `name` | string | ✓ | 룰 이름 (한글, 짧게) |
| `category` | enum | ✓ | `format` / `terminology` / `table` / `required_field` / `cross_reference` |
| `severity` | enum | ✓ | `required` (필수, 빨강) / `recommended` (권고, 주황) / `info` (참고, 파랑) |
| `check_type` | enum | ✓ | `regex_replace` / `regex_find` / `table_header` / `required_section` / `format_check` |
| `message` | string | ✓ | 검토자에게 보여줄 메시지 (한글) |
| `rationale` | string | | 왜 이 룰이 있는지 (SOP 근거, 회사 컨벤션 등) |
| `enabled` | bool | | 기본값 `true`. 끄고 싶을 때 `false` |

## check_type별 추가 필드

### `regex_replace` (잘못된 표현 → 올바른 표현)
- `pattern`: 찾을 정규식
- `replacement`: 대체할 문자열
- `example_wrong`: 잘못된 예시
- `example_right`: 올바른 예시

### `regex_find` (해당 패턴이 있으면 flag)
- `pattern`: 찾을 정규식
- `forbidden`: true면 발견 시 flag, false면 미발견 시 flag

### `table_header` (특정 표의 헤더 검증)
- `table_index`: 표 번호 (1부터)
- `expected_headers`: 기대 헤더 리스트
- `forbidden_keywords`: 헤더에 있으면 안 되는 단어
- `language`: `ko` / `en` / `bilingual`

### `required_section` (필수 섹션 존재 확인)
- `section_name`: 섹션명
- `keywords`: 섹션 식별 키워드

### `format_check` (서식 검증)
- `target`: `font_size` / `line_spacing` / `paragraph_spacing` / `alignment`
- `expected_value`: 기대값
