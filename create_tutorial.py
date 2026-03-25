"""ETF 유튜브 대시보드 강의안 Word 문서 생성 (6회차 x 30분 구성)"""

from docx import Document
from docx.shared import Pt, Inches, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
import os

doc = Document()

# ── 스타일 설정 ──
style = doc.styles['Normal']
font = style.font
font.name = '맑은 고딕'
font.size = Pt(11)
style.element.rPr.rFonts.set(qn('w:eastAsia'), '맑은 고딕')

def add_heading_styled(text, level=1):
    h = doc.add_heading(text, level=level)
    for run in h.runs:
        run.font.name = '맑은 고딕'
        run.element.rPr.rFonts.set(qn('w:eastAsia'), '맑은 고딕')
    return h

def add_para(text, bold=False, italic=False, size=None, color=None, align=None):
    p = doc.add_paragraph()
    run = p.add_run(text)
    run.font.name = '맑은 고딕'
    run.element.rPr.rFonts.set(qn('w:eastAsia'), '맑은 고딕')
    if bold:
        run.bold = True
    if italic:
        run.italic = True
    if size:
        run.font.size = Pt(size)
    if color:
        run.font.color.rgb = RGBColor(*color)
    if align:
        p.alignment = align
    return p

def add_code_block(text):
    p = doc.add_paragraph()
    run = p.add_run(text)
    run.font.name = 'Consolas'
    run.font.size = Pt(10)
    run.font.color.rgb = RGBColor(0, 100, 0)
    pf = p.paragraph_format
    pf.left_indent = Cm(1)
    pf.space_before = Pt(4)
    pf.space_after = Pt(4)
    return p

def add_error_table(rows):
    table = doc.add_table(rows=1, cols=3)
    table.style = 'Light Grid Accent 1'
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    headers = table.rows[0].cells
    headers[0].text = '증상'
    headers[1].text = '원인'
    headers[2].text = '해결법'
    for h in headers:
        for p in h.paragraphs:
            for r in p.runs:
                r.bold = True
    for symptom, cause, fix in rows:
        row = table.add_row().cells
        row[0].text = symptom
        row[1].text = cause
        row[2].text = fix
    doc.add_paragraph()

def add_checkpoint(items):
    for item in items:
        p = doc.add_paragraph()
        run = p.add_run(item)
        run.font.name = '맑은 고딕'
        run.element.rPr.rFonts.set(qn('w:eastAsia'), '맑은 고딕')
        run.font.color.rgb = RGBColor(0, 128, 0)

# ════════════════════════════════════════════════════
# 표지
# ════════════════════════════════════════════════════
doc.add_paragraph()
doc.add_paragraph()
add_para('완전 초보자를 위한', size=14, align=WD_ALIGN_PARAGRAPH.CENTER, color=(100, 100, 100))
add_para('Claude Code로 만드는\nETF 유튜브 데이터 대시보드', size=22, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER)
doc.add_paragraph()
add_para('컴퓨터를 처음 배우는 중학생도 혼자 완주할 수 있는 실습 강의안', size=12, align=WD_ALIGN_PARAGRAPH.CENTER, color=(100, 100, 100))
doc.add_paragraph()
add_para('구성: 총 4회차 (회차당 30~40분) | 총 소요 시간: 약 2시간 20분', size=12, align=WD_ALIGN_PARAGRAPH.CENTER, bold=True)
add_para('전제 조건: 인터넷 검색과 유튜브 시청이 가능하면 충분합니다!', size=11, align=WD_ALIGN_PARAGRAPH.CENTER, italic=True)
doc.add_page_break()

# ════════════════════════════════════════════════════
# 시작하기 전에
# ════════════════════════════════════════════════════
add_heading_styled('시작하기 전에 - 이 강의안 사용법', level=1)
doc.add_paragraph()
add_para('이 강의안에서 자주 나오는 표시의 뜻:', bold=True)
doc.add_paragraph()
for icon in [
    '>> 중요한 개념 설명',
    'TIP - 알아두면 편한 것',
    '!! 막혔을 때 - 문제가 생겼을 때 해결법',
    'CHECK - 여기까지 잘 따라왔는지 확인',
]:
    p = doc.add_paragraph()
    p.add_run(icon).bold = True

doc.add_paragraph()

# ── 스크린샷 도움 요청법 ──
add_para('>> 막혔을 때 스크린샷으로 Claude에게 도움 요청하는 법', bold=True, size=13)
doc.add_paragraph('실습 중 에러가 나거나 막히면, 화면을 캡처해서 Claude Code에게 보여줄 수 있어요!')
doc.add_paragraph('Claude가 화면을 보고 정확하게 도와줍니다.')
doc.add_paragraph()

doc.add_paragraph('방법 1: 스크린샷 파일을 끌어다 놓기 (가장 쉬운 방법)')
doc.add_paragraph('  1) 키보드에서 "윈도우 키 + Shift + S"를 동시에 누르세요')
doc.add_paragraph('  2) 화면이 어두워지면, 캡처하고 싶은 부분을 마우스로 드래그하세요')
doc.add_paragraph('  3) 오른쪽 아래 알림을 클릭해서 바탕화면에 저장하세요')
doc.add_paragraph('  4) 저장된 스크린샷 파일을 Claude Code 터미널 창으로 끌어다 놓으세요')
doc.add_paragraph('  5) 파일 경로가 자동 입력되면, 앞에 질문을 추가하세요:')
add_code_block('이 에러 화면 좀 봐줘 /Users/이름/Desktop/screenshot.png')
doc.add_paragraph()

doc.add_paragraph('방법 2: 클립보드에서 바로 붙여넣기')
doc.add_paragraph('  1) "윈도우 키 + Shift + S"로 화면 캡처 (클립보드에 자동 복사)')
doc.add_paragraph('  2) Claude Code에서 Ctrl+V로 붙여넣기')
doc.add_paragraph('  3) 질문과 함께 Enter!')
doc.add_paragraph()

doc.add_paragraph('방법 3: 파일 경로 직접 입력')
add_code_block('이 화면에서 뭐가 잘못된 건지 알려줘 ~/Desktop/캡처.png')
doc.add_paragraph()
doc.add_paragraph('TIP: 에러 메시지가 전부 보이게 넓게 캡처하세요!')
doc.add_paragraph('TIP: "이 에러 어떻게 해결해?" 같은 질문을 함께 적으면 더 정확한 답을 받아요.')
doc.add_paragraph()

# ── 용어 미리보기 ──
add_para('>> 용어 미리보기', bold=True, size=13)
terms = [
    ('터미널', '컴퓨터에게 글자로 명령을 내리는 채팅창. 카카오톡처럼 내가 글을 치면 컴퓨터가 답해줍니다.'),
    ('코드', '컴퓨터가 알아듣는 언어로 쓴 편지. 우리가 한국어를 쓰듯, 컴퓨터에게는 코드로 말합니다.'),
    ('명령 프롬프트(cmd)', '윈도우에 옛날부터 있던 가장 기본 터미널. 기능이 적어서 요즘은 잘 안 써요.'),
    ('PowerShell', '윈도우에 기본으로 있는 고급 터미널. 명령 프롬프트의 업그레이드 버전이에요.'),
    ('Git', '코드의 타임머신. 작업 내용을 저장해두면, 언제든 과거로 돌아가거나 다른 사람과 공유할 수 있어요.'),
    ('Git Bash', 'Git과 함께 설치되는 특별한 터미널. 전 세계 개발자들이 쓰는 명령어를 윈도우에서도 쓸 수 있게 해줘요.'),
    ('GitHub', '코드의 인스타그램. 내 코드를 올려서 보관하고 다른 사람과 공유하는 웹사이트.'),
    ('API', '앱과 앱 사이의 택배 시스템. 유튜브한테 "이 영상 정보 좀 줘"라고 부탁하는 우체통 같은 것.'),
    ('대시보드', '여러 정보를 한 화면에 보기 좋게 정리한 게시판. 자동차 계기판처럼 중요한 숫자를 한눈에 보여줍니다.'),
    ('ETF', '여러 회사의 주식을 한 바구니에 담아 파는 묶음 상품. 과자 종합선물세트 같은 것!'),
]
for term, desc in terms:
    p = doc.add_paragraph()
    run1 = p.add_run(f'  {term}: ')
    run1.bold = True
    p.add_run(desc)

doc.add_page_break()

# ════════════════════════════════════════════════════════════
# ██  1회차: 터미널부터 AI 비서까지 한 번에!
# ════════════════════════════════════════════════════════════
add_heading_styled('1회차: 터미널부터 AI 비서까지 한 번에!', level=1)
add_para('소요 시간: 40분', bold=True, color=(0, 100, 200))
add_para('이 회차를 마치면: 터미널, Git, Node.js를 설치하고, AI 비서 Claude Code와 첫 대화까지! "와, 이런 거구나!" 체험을 바로 해볼 수 있어요.', bold=True)
doc.add_paragraph()

add_heading_styled('Step 1. Microsoft Store에서 Windows Terminal 설치 (5분)', level=2)
doc.add_paragraph('키보드에서 윈도우 키(키보드 왼쪽 아래 창문 모양)를 한 번 누르세요.')
doc.add_paragraph('검색창이 뜨면 아래 글자를 그대로 입력하세요:')
add_code_block('Microsoft Store')
doc.add_paragraph('파란 쇼핑백 모양 아이콘이 보이면 클릭하세요.')
doc.add_paragraph()
doc.add_paragraph('Microsoft Store가 열리면, 위쪽 검색창에 입력하세요:')
add_code_block('Windows Terminal')
doc.add_paragraph('"Windows Terminal" 앱을 찾아 "설치" 또는 "다운로드" 버튼을 누르세요.')
doc.add_paragraph('TIP: 이미 "열기" 버튼이 보이면 이미 설치된 거예요! 바로 다음 Step으로!')
doc.add_paragraph()

add_heading_styled('Step 2. 터미널 처음 열어보기 (3분)', level=2)
doc.add_paragraph('설치 후 "열기" 버튼을 누르세요. (또는 윈도우 키 -> "Terminal" 검색)')
doc.add_paragraph('까만(또는 파란) 창이 뜹니다. 이게 바로 터미널이에요!')
doc.add_paragraph()
doc.add_paragraph('>> 무서워하지 마세요! 이건 컴퓨터랑 카톡하는 창이에요.')
doc.add_paragraph('>> 내가 글을 쓰면 컴퓨터가 대답합니다.')
doc.add_paragraph()
doc.add_paragraph('화면에 이런 비슷한 글자가 보일 거예요:')
add_code_block('PS C:\\Users\\여러분이름>')
doc.add_paragraph('"나는 준비됐어, 명령을 내려줘!"라는 뜻이에요.')
doc.add_paragraph()

add_heading_styled('Step 3. 첫 번째 명령어 입력해보기 (2분)', level=2)
doc.add_paragraph('아래 명령어를 정확히 따라 입력하고 Enter 키를 누르세요:')
add_code_block('echo "안녕, 나는 터미널이야!"')
doc.add_paragraph('화면에 이렇게 나와야 해요:')
add_code_block('안녕, 나는 터미널이야!')
doc.add_paragraph('축하해요! 방금 컴퓨터에게 첫 번째 명령을 내렸습니다!')
doc.add_paragraph()

add_heading_styled('Step 4. 터미널의 종류 알아보기 (5분)', level=2)
doc.add_paragraph('>> 잠깐! 윈도우에는 터미널(까만 창)이 여러 종류가 있어요.')
doc.add_paragraph('>> 처음에는 헷갈리지만, 비유로 쉽게 이해해봅시다!')
doc.add_paragraph()

doc.add_paragraph('1) 명령 프롬프트(cmd) = "일반 자전거"')
doc.add_paragraph('   - 윈도우에 처음부터 있는 가장 오래된 터미널')
doc.add_paragraph('   - 할 수 있는 것이 적고, 요즘 개발에는 잘 안 써요')
doc.add_paragraph('   - 화면에 C:\\Users\\이름> 이렇게 보여요')
doc.add_paragraph()

doc.add_paragraph('2) PowerShell(파워셸) = "전동 자전거"')
doc.add_paragraph('   - cmd를 업그레이드한 버전. 윈도우 10/11에 기본 설치')
doc.add_paragraph('   - cmd보다 많은 기능이 있지만, 개발자용 도구가 부족할 때 있어요')
doc.add_paragraph('   - 화면에 PS C:\\Users\\이름> 이렇게 보여요 (앞에 PS가 붙음)')
doc.add_paragraph('   - 방금 우리가 연 Windows Terminal이 기본으로 PowerShell을 보여줘요')
doc.add_paragraph()

doc.add_paragraph('3) Git Bash = "전동 킥보드" (우리가 앞으로 쓸 것!)')
doc.add_paragraph('   - Git을 설치하면 함께 오는 특별한 터미널')
doc.add_paragraph('   - 전 세계 개발자들이 쓰는 리눅스/맥 스타일 명령어를 윈도우에서도 사용 가능')
doc.add_paragraph('   - Claude Code가 가장 잘 동작하는 환경!')
doc.add_paragraph('   - 화면에 이름@컴퓨터 MINGW64 ~ 이렇게 보이고, $ 표시로 명령을 기다려요')
doc.add_paragraph()
doc.add_paragraph('>> 결론: 다음 회차에서 Git Bash를 설치할 거예요!')
doc.add_paragraph('>> 지금 연 PowerShell은 "이런 게 있구나" 정도로만 알아두면 돼요.')
doc.add_paragraph()

add_heading_styled('Step 5. Git 설치하기 (10분)', level=2)
doc.add_paragraph()
doc.add_paragraph('>> Git을 왜 설치해야 하나요?')
doc.add_paragraph('Git은 개발자에게 필수인 "코드 타임머신"이에요. 3가지 이유:')
doc.add_paragraph()
doc.add_paragraph('  1) 실수 복구: 코드를 잘못 바꿨을 때 "아까 그 상태로 돌려줘!"가 가능')
doc.add_paragraph('     비유: 게임의 "세이브 포인트". 중간중간 저장해두면 실패해도 다시 시작!')
doc.add_paragraph()
doc.add_paragraph('  2) 협업: 여러 사람이 같은 프로젝트에서 동시에 작업 가능')
doc.add_paragraph('     비유: 구글 문서에서 여러 명이 동시에 편집하는 것과 비슷')
doc.add_paragraph()
doc.add_paragraph('  3) Git Bash: Git을 설치하면 "Git Bash"라는 특별한 터미널이 함께 설치!')
doc.add_paragraph('     Claude Code는 이 Git Bash에서 가장 잘 동작합니다')
doc.add_paragraph()

doc.add_paragraph('브라우저에서 아래 주소로 이동하세요:')
add_code_block('https://git-scm.com/downloads/win')
doc.add_paragraph('"Click here to download" 또는 "64-bit Git for Windows Setup"을 클릭하세요.')
doc.add_paragraph()

doc.add_paragraph('다운로드된 파일(Git-숫자-64-bit.exe)을 더블클릭하세요.')
doc.add_paragraph('"이 앱이 변경하도록 허용하시겠습니까?" -> "예"')
doc.add_paragraph()
doc.add_paragraph('설치 화면이 나오면:')
doc.add_paragraph('  1) "Next" 버튼을 계속 누르세요 (약 8~10번)')
doc.add_paragraph('  2) 중간에 여러 옵션이 나오지만, 기본값 그대로 두고 "Next"만!')
doc.add_paragraph('  3) 마지막에 "Install" 버튼 클릭')
doc.add_paragraph('  4) 설치 끝나면 "Finish" 클릭')
doc.add_paragraph()
doc.add_paragraph('TIP: 영어 옵션이 많지만 걱정 마세요!')
doc.add_paragraph('TIP: "Next -> Next -> ... -> Install -> Finish" 순서만 기억!')
doc.add_paragraph()

add_heading_styled('Step 6. Git Bash 열고 확인하기 (5분)', level=2)
doc.add_paragraph('윈도우 키 -> "Git Bash" 검색 -> 클릭')
doc.add_paragraph()
doc.add_paragraph('까만 창이 열리면서 이런 글자가 보여요:')
add_code_block('사용자이름@컴퓨터이름 MINGW64 ~\n$')
doc.add_paragraph('$ 표시가 "명령을 기다리고 있어요"라는 뜻!')
doc.add_paragraph()
doc.add_paragraph('>> 앞으로 이 강의에서 "터미널을 열어주세요"라고 하면 Git Bash를 열어주세요!')
doc.add_paragraph()

doc.add_paragraph('Git이 잘 설치됐는지 확인해봅시다:')
add_code_block('git --version')
doc.add_paragraph('이런 결과가 나오면 성공:')
add_code_block('git version 2.47.1.windows.2')
doc.add_paragraph('(숫자는 달라도 OK. 숫자가 나오기만 하면 성공!)')
doc.add_paragraph()

doc.add_paragraph('프로젝트 폴더도 미리 만들어둡시다:')
add_code_block('mkdir ~/etf-dashboard')
doc.add_paragraph('>> mkdir = "make directory" = "폴더를 만들어줘"')
doc.add_paragraph()

add_heading_styled('Step 7. Node.js 설치하기 (7분)', level=2)
doc.add_paragraph('>> Node.js(노드제이에스)가 뭔가요?')
doc.add_paragraph('프로그램을 실행시켜주는 "엔진"이에요.')
doc.add_paragraph('비유: 자동차(프로그램)가 달리려면 엔진(Node.js)이 필요한 것처럼,')
doc.add_paragraph('Claude Code를 돌리려면 Node.js가 먼저 필요해요.')
doc.add_paragraph()

doc.add_paragraph('브라우저에서 아래 주소로 이동하세요:')
add_code_block('https://nodejs.org')
doc.add_paragraph('"LTS"라고 적힌 큰 초록색 버튼을 클릭하세요.')
doc.add_paragraph('>> LTS = "Long Term Support" = "오래오래 안정적으로 쓸 수 있는 버전". 항상 이걸 선택!')
doc.add_paragraph()

doc.add_paragraph('다운로드된 파일(node-v숫자-x64.msi)을 더블클릭하세요.')
doc.add_paragraph('"Next -> Next -> ... -> Install -> Finish" 순서로 설치!')
doc.add_paragraph()

doc.add_paragraph('!! 중요: 설치 후 Git Bash를 껐다가 다시 열어야 합니다!')
doc.add_paragraph('Git Bash를 새로 열고 확인:')
add_code_block('node --version')
doc.add_paragraph('v로 시작하는 숫자가 나오면 성공! (예: v22.13.1)')
doc.add_paragraph()

add_heading_styled('Step 8. Claude Code 설치 + 첫 실행! (8분)', level=2)
doc.add_paragraph('>> Claude Code가 뭔가요?')
doc.add_paragraph('터미널 안에서 동작하는 AI 비서예요.')
doc.add_paragraph('비유: 카카오톡으로 똑똑한 친구에게 "이거 만들어줘"라고 부탁하는 것처럼,')
doc.add_paragraph('터미널에서 Claude에게 프로그램을 만들어달라고 부탁할 수 있어요.')
doc.add_paragraph()

doc.add_paragraph('Git Bash에서 아래 명령어를 입력하세요:')
add_code_block('npm install -g @anthropic-ai/claude-code')
doc.add_paragraph('>> npm install = "설치해줘", -g = "어디서든 쓸 수 있게"')
doc.add_paragraph('설치에 1~3분 걸려요. "added XX packages"가 보이면 성공!')
doc.add_paragraph()

doc.add_paragraph('프로젝트 폴더로 이동하고 Claude를 실행합시다:')
add_code_block('cd ~/etf-dashboard\nclaude')
doc.add_paragraph()

doc.add_paragraph('처음 실행하면 로그인이 필요합니다. 화면에 이런 메시지가 나타나요:')
add_code_block('? How would you like to authenticate? (Use arrow keys)\n> Anthropic Console (OAuth - recommended)\n  API Key')
doc.add_paragraph()
doc.add_paragraph('키보드 화살표로 "Anthropic Console"이 선택된 상태에서 Enter를 누르세요.')
doc.add_paragraph('브라우저가 자동으로 열려요.')
doc.add_paragraph()

doc.add_paragraph('브라우저에서:')
doc.add_paragraph('  - 계정이 있다면: 로그인')
doc.add_paragraph('  - 계정이 없다면: "Sign up" 클릭 -> 구글 계정으로 간편 가입 가능!')
doc.add_paragraph('  - 로그인 후 "Allow" 또는 "허용" 클릭')
doc.add_paragraph()

doc.add_paragraph('터미널에 이런 화면이 보이면 성공:')
add_code_block('> Welcome to Claude Code!\n\n>')
doc.add_paragraph()

add_heading_styled('Step 9. Claude에게 첫 인사하기! (3분)', level=2)
doc.add_paragraph('> 표시 옆에 아래처럼 입력하고 Enter:')
add_code_block('안녕! 나는 ETF 대시보드를 만들고 싶어. 도와줄 수 있어?')
doc.add_paragraph('Claude가 한국어로 친절하게 대답해줄 거예요!')
doc.add_paragraph()
doc.add_paragraph('축하해요! 오늘 한 번에 터미널, Git, Node.js, AI 비서까지 만났습니다!')
doc.add_paragraph()
doc.add_paragraph('TIP: Claude Code를 끝내려면 Ctrl+C 또는 /exit 입력')
doc.add_paragraph()

add_heading_styled('CHECK 체크포인트', level=2)
add_checkpoint([
    'CHECK: Windows Terminal이 설치되었나요?',
    'CHECK: echo 명령어로 글자가 화면에 나왔나요?',
    'CHECK: Git Bash 창이 열려 있나요?',
    'CHECK: git --version을 쳤을 때 숫자가 나왔나요?',
    'CHECK: node --version을 쳤을 때 숫자가 나왔나요?',
    'CHECK: Claude Code에서 "Welcome to Claude Code" 화면이 나왔나요?',
    'CHECK: Claude에게 메시지를 보내고 답변을 받았나요?',
])
doc.add_paragraph()

add_heading_styled('!! 막혔을 때', level=2)
add_error_table([
    ('Microsoft Store가 안 열려요', '윈도우 업데이트 필요', '윈도우 키 -> "업데이트 확인" 검색 -> 업데이트 실행 후 재시도'),
    ('글자를 쳤는데 반응이 없어요', '커서가 터미널 밖에 있음', '터미널 까만 화면 안쪽을 마우스로 한 번 클릭'),
    ('Git Bash가 검색해도 안 나와요', 'Git 설치 안 됨', 'Git 설치 파일을 다시 다운로드하고 실행'),
    ('설치 중 옵션이 너무 많아서 헷갈려요', '기본값이면 충분', '모든 화면에서 아무것도 바꾸지 말고 "Next"만 누르세요'),
    ('"node" 인식할 수 없는 명령', 'Git Bash를 새로 안 열었어요', 'Git Bash를 완전히 닫고 새로 여세요'),
    ('"claude" 인식할 수 없는 명령', 'npm 설치 실패', 'Git Bash 닫고 다시 열기. 안 되면 npm install -g @anthropic-ai/claude-code 다시 실행'),
    ('로그인 페이지가 안 열려요', '브라우저 팝업 차단', '터미널에 나온 URL을 복사해서 브라우저 주소창에 직접 붙여넣기'),
    ('영어로 대답해요', '언어 자동 감지 실패', '"한국어로 대답해줘"라고 입력'),
])

add_para('1회차 완료! 수고하셨습니다!', bold=True, size=12, color=(0, 128, 0))
add_para('오늘 한 번에 터미널, Git, Node.js, Claude Code까지 설치했어요! 대단해요!', italic=True)
add_para('다음 회차에서 Python을 설치하고 유튜브 API 열쇠를 받아요.', italic=True)
doc.add_page_break()

# ════════════════════════════════════════════════════════════
# ██  2회차: Python 설치 + YouTube API 키 발급
# ════════════════════════════════════════════════════════════
add_heading_styled('2회차: Python 설치 + YouTube API 키 발급', level=1)
add_para('소요 시간: 35분', bold=True, color=(0, 100, 200))
add_para('이 회차를 마치면: Python이 설치되고, 유튜브에서 데이터를 가져올 수 있는 열쇠(API 키)를 갖게 되고, 데이터 수집 프로그램까지 만들어져요!', bold=True)
doc.add_paragraph()

add_heading_styled('Step 1. Python 설치하기 (10분)', level=2)
doc.add_paragraph('>> Python(파이썬)이 뭔가요?')
doc.add_paragraph('컴퓨터 프로그래밍 언어예요. 영어처럼 읽기 쉬워서 초보자에게 인기가 많아요.')
doc.add_paragraph('우리는 데이터를 수집할 때 Python을 사용할 거예요.')
doc.add_paragraph()

doc.add_paragraph('브라우저에서 아래 주소로 이동하세요:')
add_code_block('https://www.python.org/downloads/')
doc.add_paragraph('노란색 "Download Python 3.x.x" 버튼을 클릭하세요.')
doc.add_paragraph()

doc.add_paragraph('다운로드된 파일을 실행하면 설치 화면이 나옵니다.')
doc.add_paragraph()
doc.add_paragraph('!! 가장 중요한 순간!')
doc.add_paragraph('!! 맨 아래에 "Add python.exe to PATH"라는 체크박스가 보입니다.')
doc.add_paragraph('!! 이것을 반.드.시. 체크하세요!')
doc.add_paragraph('!! (체크 안 하면 나중에 python 명령어가 안 먹어서 큰 문제가 생겨요)')
doc.add_paragraph()
doc.add_paragraph('체크한 후 "Install Now"를 클릭하세요.')
doc.add_paragraph('설치가 끝나면 "Close"를 누르세요.')
doc.add_paragraph()

doc.add_paragraph('Git Bash를 껐다가 다시 열고 확인:')
add_code_block('python --version')
doc.add_paragraph('Python 3.x.x가 나오면 성공!')
doc.add_paragraph()
add_code_block('pip --version')
doc.add_paragraph('>> pip = Python의 앱스토어. 필요한 재료를 설치하는 도구. 숫자가 나오면 OK!')
doc.add_paragraph()

add_heading_styled('Step 2. Google Cloud Console 접속 (3분)', level=2)
doc.add_paragraph('>> API 키가 뭔가요?')
doc.add_paragraph('"출입증"이에요. 놀이공원에 입장권이 필요하듯,')
doc.add_paragraph('유튜브 데이터를 가져오려면 구글이 발급해주는 출입증(API 키)이 필요합니다.')
doc.add_paragraph('이 출입증은 무료예요! (하루 사용 횟수 제한만 있어요)')
doc.add_paragraph()

doc.add_paragraph('브라우저에서 아래 주소로 이동하세요:')
add_code_block('https://console.cloud.google.com')
doc.add_paragraph('구글 계정으로 로그인하세요. (지메일 있으면 OK!)')
doc.add_paragraph()

add_heading_styled('Step 3. 새 프로젝트 만들기 (3분)', level=2)
doc.add_paragraph('화면 왼쪽 위 "프로젝트 선택" (또는 "Select a project") 클릭')
doc.add_paragraph('"새 프로젝트" (또는 "New Project") 클릭')
doc.add_paragraph()
doc.add_paragraph('프로젝트 이름에 아래처럼 입력:')
add_code_block('etf-youtube-dashboard')
doc.add_paragraph('"만들기" (또는 "Create") 클릭')
doc.add_paragraph('잠시 기다리면 프로젝트가 만들어집니다.')
doc.add_paragraph()

add_heading_styled('Step 4. YouTube Data API 활성화 (3분)', level=2)
doc.add_paragraph('왼쪽 메뉴: "API 및 서비스" -> "라이브러리"')
doc.add_paragraph('(영어: "APIs & Services" -> "Library")')
doc.add_paragraph()
doc.add_paragraph('검색창에 입력:')
add_code_block('YouTube Data API v3')
doc.add_paragraph('"YouTube Data API v3" 클릭 -> 파란색 "사용" (Enable) 버튼 클릭')
doc.add_paragraph('>> 이 단계는 "유튜브 데이터 택배 서비스를 신청합니다"라는 뜻!')
doc.add_paragraph()

add_heading_styled('Step 5. API 키 만들기 (3분)', level=2)
doc.add_paragraph('왼쪽 메뉴: "API 및 서비스" -> "사용자 인증 정보"')
doc.add_paragraph('(영어: "APIs & Services" -> "Credentials")')
doc.add_paragraph()
doc.add_paragraph('"+ 사용자 인증 정보 만들기" -> "API 키" 클릭')
doc.add_paragraph('(영어: "Create Credentials" -> "API key")')
doc.add_paragraph()
doc.add_paragraph('API 키가 생성됩니다! 이런 모양의 긴 글자:')
add_code_block('AIzaSyD-xxxxxxxxxxxxxxxxxxxxxxxxx')
doc.add_paragraph()
doc.add_paragraph('!! 매우 중요: 이 키를 복사해서 메모장에 저장하세요!')
doc.add_paragraph('!! 비밀번호처럼 남에게 보여주면 안 됩니다!')
doc.add_paragraph()

add_heading_styled('Step 6. API 키를 프로젝트에 저장 (3분)', level=2)
doc.add_paragraph('Git Bash에서:')
add_code_block('cd ~/etf-dashboard\necho "YOUTUBE_API_KEY=여러분의_API_키_여기에" > .env')
doc.add_paragraph('("여러분의_API_키_여기에" 부분에 아까 복사한 키를 붙여넣으세요)')
doc.add_paragraph()
doc.add_paragraph('>> .env 파일은 "비밀 메모장"이에요.')
doc.add_paragraph('>> 프로그램이 필요할 때 여기서 API 키를 꺼내 씁니다.')
doc.add_paragraph()
doc.add_paragraph('잘 저장됐는지 확인:')
add_code_block('cat .env')
doc.add_paragraph('결과:')
add_code_block('YOUTUBE_API_KEY=AIzaSyD-xxxxxxxxxxxxxxxxxxxxxxxxx')
doc.add_paragraph()

add_heading_styled('Step 7. Claude에게 데이터 수집 프로그램 만들어달라고 부탁 (5분)', level=2)
doc.add_paragraph('>> 이번 Step의 핵심: 우리가 직접 코드를 쓰지 않아요!')
doc.add_paragraph('>> Claude AI에게 "이런 프로그램 만들어줘"라고 말하면, AI가 대신 만들어줍니다.')
doc.add_paragraph()

doc.add_paragraph('Claude Code를 실행하세요:')
add_code_block('claude')
doc.add_paragraph()
doc.add_paragraph('Claude의 > 프롬프트에 아래 내용을 통째로 복사해서 붙여넣으세요:')
add_code_block(
    '.env 파일에 YOUTUBE_API_KEY가 저장되어 있어.\n'
    '이 키를 사용해서 유튜브에서 ETF 관련 영상을 검색하고\n'
    '데이터를 수집하는 Python 프로그램을 만들어줘.\n'
    '\n'
    '요구사항:\n'
    '1. "ETF 투자", "ETF 추천", "ETF 분석" 키워드로 검색\n'
    '2. 각 키워드당 최신 영상 10개씩 수집\n'
    '3. 수집할 정보: 영상 제목, 채널명, 조회수, 좋아요 수,\n'
    '   게시일, 영상 링크\n'
    '4. 결과를 etf_youtube_data.json 파일로 저장\n'
    '5. python-dotenv를 사용해서 .env에서 API 키 읽기\n'
    '6. 실행하면 수집 진행 상황을 보여주기'
)
doc.add_paragraph()
doc.add_paragraph('Enter를 누르면 Claude가 코드를 작성하기 시작합니다!')
doc.add_paragraph('Claude가 "파일을 만들어도 될까요?" 같은 질문을 하면 "y" 입력 후 Enter')
doc.add_paragraph()

doc.add_paragraph('Claude가 작업을 마치면 /exit으로 나온 후 확인:')
add_code_block('/exit\nls')
doc.add_paragraph('이런 파일들이 보여야 해요:')
add_code_block('.env\ncollect_data.py\nrequirements.txt')
doc.add_paragraph()

add_heading_styled('CHECK 체크포인트', level=2)
add_checkpoint([
    'CHECK: python --version 결과에 Python 3.x.x가 나왔나요?',
    'CHECK: YouTube Data API v3를 활성화(Enable) 했나요?',
    'CHECK: cat .env 결과에 YOUTUBE_API_KEY=... 형태가 보이나요?',
    'CHECK: collect_data.py 파일이 만들어졌나요?',
])
doc.add_paragraph()

add_heading_styled('!! 막혔을 때', level=2)
add_error_table([
    ('python 대신 Microsoft Store가 열려요', 'Windows 앱 별칭 문제', 'Git Bash에서는 보통 안 생겨요. 생기면 python3으로 시도'),
    ('"Add python.exe to PATH" 체크를 깜빡했어요', 'PATH 미등록', 'Python을 삭제 후 다시 설치하면서 이번엔 꼭 체크!'),
    ('"결제 계정 설정하세요" 메시지', '구글 결제 정보 요청', '무료 체험 등록 - 실제 과금 안 돼요. 카드 등록은 본인 확인 용도'),
    ('API 키가 안 보여요', '팝업이 빨리 닫힘', '"사용자 인증 정보" 페이지에서 이미 만든 키 확인 가능'),
    ('"프로젝트 선택"이 안 보여요', '화면 상단 확인', '"Google Cloud" 로고 옆 드롭다운 메뉴를 클릭'),
    ('Claude가 다른 파일명을 썼어요', 'Claude가 때때로 다른 이름 선택', '괜찮아요! ls로 확인하고 Claude가 만든 이름으로 진행'),
])

add_para('2회차 완료! 수고하셨습니다!', bold=True, size=12, color=(0, 128, 0))
add_para('다음 회차에서 프로그램을 실행해서 진짜 데이터를 모아올 거예요!', italic=True)
doc.add_page_break()

# ════════════════════════════════════════════════════════════
# ██  3회차: 데이터 수집 실행 + 대시보드 제작
# ════════════════════════════════════════════════════════════
add_heading_styled('3회차: 데이터 수집 실행 + 대시보드 제작', level=1)
add_para('소요 시간: 35분', bold=True, color=(0, 100, 200))
add_para('이 회차를 마치면: 유튜브에서 ETF 데이터를 진짜로 모아오고, 그 데이터를 예쁜 대시보드로 만들어 브라우저에서 볼 수 있어요! 핵심 프로젝트 완성!', bold=True)
doc.add_paragraph()

add_heading_styled('Step 1. 필요한 재료(라이브러리) 설치 (3분)', level=2)
doc.add_paragraph('Git Bash에서 프로젝트 폴더로 이동하세요:')
add_code_block('cd ~/etf-dashboard')
doc.add_paragraph()
doc.add_paragraph('프로그램에 필요한 재료를 설치합니다:')
add_code_block('pip install -r requirements.txt')
doc.add_paragraph('>> "requirements.txt에 적힌 재료들을 전부 설치해줘"라는 뜻')
doc.add_paragraph('설치에 30초~1분 걸려요.')
doc.add_paragraph()

add_heading_styled('Step 2. 데이터 수집 프로그램 실행! (5분)', level=2)
doc.add_paragraph('드디어! 프로그램을 실행해봅시다:')
add_code_block('python collect_data.py')
doc.add_paragraph()
doc.add_paragraph('화면에 이런 진행 상황이 나타날 거예요:')
add_code_block(
    '"ETF 투자" 검색 중...\n'
    '  10개 영상 수집 완료\n'
    '"ETF 추천" 검색 중...\n'
    '  10개 영상 수집 완료\n'
    '"ETF 분석" 검색 중...\n'
    '  10개 영상 수집 완료\n'
    '\n'
    '총 30개 영상 데이터 수집 완료!\n'
    'etf_youtube_data.json에 저장했습니다.'
)
doc.add_paragraph()
doc.add_paragraph('축하해요! 유튜브에서 ETF 데이터를 자동으로 모아왔습니다!')
doc.add_paragraph()

doc.add_paragraph('어떤 데이터가 모였는지 살짝 확인해볼까요?')
add_code_block("python -c \"import json; data=json.load(open('etf_youtube_data.json',encoding='utf-8')); print(f'총 {len(data)}개 영상'); print(data[0]['title'])\"")
doc.add_paragraph('첫 번째 영상의 제목이 화면에 나와요!')
doc.add_paragraph()

add_heading_styled('Step 3. Claude에게 대시보드 만들어달라고 부탁 (12분)', level=2)
doc.add_paragraph('>> 대시보드가 뭔가요?')
doc.add_paragraph('자동차 계기판(dashboard)처럼, 중요한 숫자와 그래프를 한 화면에 보여주는 웹 페이지!')
doc.add_paragraph()

doc.add_paragraph('Claude Code를 실행하세요:')
add_code_block('claude')
doc.add_paragraph()
doc.add_paragraph('아래 내용을 통째로 복사해서 붙여넣으세요:')
add_code_block(
    'etf_youtube_data.json 파일에 유튜브 ETF 영상 데이터가 있어.\n'
    '이 데이터를 보여주는 예쁜 웹 대시보드를 HTML로 만들어줘.\n'
    '\n'
    '요구사항:\n'
    '1. 하나의 index.html 파일로 만들어줘 (서버 없이 바로 열 수 있게)\n'
    '2. 한국어로 작성\n'
    '3. 반드시 포함할 내용:\n'
    '   - 상단에 "ETF 유튜브 트렌드 대시보드" 제목\n'
    '   - 전체 영상 수, 총 조회수, 평균 좋아요 수 요약 카드\n'
    '   - 조회수 TOP 10 영상 막대 그래프 (Chart.js 사용)\n'
    '   - 채널별 영상 수 파이 차트\n'
    '   - 전체 영상 목록 테이블 (제목 클릭하면 유튜브로 이동)\n'
    '4. 디자인: 깔끔한 다크 테마, 카드형 레이아웃\n'
    '5. JSON 데이터를 HTML 파일 안에 직접 포함시켜줘\n'
    '6. 모바일에서도 보기 좋게 반응형으로 만들어줘'
)
doc.add_paragraph()
doc.add_paragraph('Enter! Claude가 대시보드를 만들기 시작합니다.')
doc.add_paragraph('파일 만들어도 되는지 물으면 "y" 입력')
doc.add_paragraph('코드 작성에 1~3분 걸릴 수 있어요. 기다려주세요!')
doc.add_paragraph()

add_heading_styled('Step 4. 브라우저에서 대시보드 열기! (5분)', level=2)
doc.add_paragraph('Claude가 작업을 마치면:')
add_code_block('/exit\nls')
doc.add_paragraph('index.html 파일이 보이면 성공!')
doc.add_paragraph()

doc.add_paragraph('이 순간을 위해 여기까지 왔습니다!')
add_code_block('start index.html')
doc.add_paragraph()
doc.add_paragraph('인터넷 브라우저가 열리면서 여러분이 만든 대시보드가 나타납니다!')
doc.add_paragraph()
doc.add_paragraph('화면에 이런 것들이 보여요:')
doc.add_paragraph('  - 맨 위: "ETF 유튜브 트렌드 대시보드" 제목')
doc.add_paragraph('  - 요약 카드: 전체 영상 수, 총 조회수, 평균 좋아요 수')
doc.add_paragraph('  - 막대 그래프: 조회수 TOP 10 영상')
doc.add_paragraph('  - 파이 차트: 채널별 영상 수')
doc.add_paragraph('  - 영상 목록 테이블: 제목 클릭하면 유튜브로 이동!')
doc.add_paragraph()

add_heading_styled('Step 5. 대시보드 꾸미기 - 도전 과제! (5분)', level=2)
doc.add_paragraph('마음에 안 드는 부분이 있다면? Claude에게 수정을 부탁하세요!')
add_code_block(
    'claude\n'
    '> index.html의 배경색을 좀 더 밝게 바꿔줘\n'
    '> 그래프 색상을 파란색 계열로 바꿔줘\n'
    '> 조회수 순 정렬 기능을 추가해줘'
)
doc.add_paragraph('수정 후 브라우저에서 Ctrl+R(새로고침)을 누르면 변경 확인!')
doc.add_paragraph()

add_heading_styled('CHECK 체크포인트', level=2)
add_checkpoint([
    'CHECK: python collect_data.py 실행 시 "수집 완료" 메시지가 나왔나요?',
    'CHECK: etf_youtube_data.json 파일이 만들어졌나요?',
    'CHECK: index.html 파일이 만들어졌나요?',
    'CHECK: 브라우저에서 대시보드가 열렸나요?',
    'CHECK: 그래프와 표가 화면에 보이나요?',
])
doc.add_paragraph()

add_heading_styled('!! 막혔을 때', level=2)
add_error_table([
    ('ModuleNotFoundError', '재료 미설치', 'pip install -r requirements.txt 다시 실행'),
    ('API key invalid / 403 에러', 'API 키 문제', '.env 파일 키 재확인 + Google Cloud에서 API "사용" 상태 확인'),
    ('인코딩 에러 (UnicodeDecodeError)', '한글 처리 문제', 'Claude에게 "인코딩 에러 나요. utf-8로 수정해줘" 부탁'),
    ('브라우저에 아무것도 안 보여요', 'HTML 파일 문제', 'Claude에게 "index.html이 안 열려요. 확인해줘" 부탁'),
    ('그래프가 안 보이고 빈 화면', '데이터 미포함', 'Claude에게 "그래프 안 보여요. 데이터 다시 확인하고 HTML에 넣어줘" 부탁'),
    ('start 명령이 안 먹어요', '명령어 문제', 'etf-dashboard 폴더에서 index.html 파일을 직접 더블클릭'),
])

add_para('3회차 완료! 핵심 프로젝트가 완성됐어요!', bold=True, size=12, color=(0, 128, 0))
add_para('마지막 회차에서 GitHub에 올려서 세상에 공유해봅시다!', italic=True)
doc.add_page_break()

# ════════════════════════════════════════════════════════════
# ██  4회차: GitHub에 공유하기
# ════════════════════════════════════════════════════════════
add_heading_styled('4회차: GitHub에 내 프로젝트 공유하기', level=1)
add_para('소요 시간: 30분', bold=True, color=(0, 100, 200))
add_para('이 회차를 마치면: 내가 만든 프로젝트가 GitHub 웹사이트에 올라가서, 주소만 알려주면 누구나 볼 수 있어요! 전체 강의 완성!', bold=True)
doc.add_paragraph()

add_heading_styled('Step 1. GitHub가 뭔지 알아보기 (3분)', level=2)
doc.add_paragraph('>> GitHub(깃허브)는 "코드의 인스타그램"이에요!')
doc.add_paragraph('  - 인스타그램에 사진을 올리듯, GitHub에 내 코드를 올려서 보관하고 공유')
doc.add_paragraph('  - 전 세계 개발자들이 자기 프로젝트를 GitHub에 올려요')
doc.add_paragraph('  - 내 컴퓨터가 고장 나도 코드가 안전하게 보관됨')
doc.add_paragraph('  - 무료!')
doc.add_paragraph()

doc.add_paragraph('>> Repository(저장소)가 뭔가요?')
doc.add_paragraph('Repository(레포지토리, 줄여서 "레포")는 "프로젝트 폴더"예요.')
doc.add_paragraph('우리의 etf-dashboard 폴더를 GitHub에 올리면, 그게 하나의 Repository!')
doc.add_paragraph('비유: 인스타그램의 "게시물 하나" = GitHub의 "Repository 하나"')
doc.add_paragraph()

add_heading_styled('Step 2. GitHub 가입하기 (7분)', level=2)
doc.add_paragraph('브라우저에서 아래 주소로 이동:')
add_code_block('https://github.com')
doc.add_paragraph()

doc.add_paragraph('오른쪽 위 "Sign up" 버튼을 클릭하세요.')
doc.add_paragraph()
doc.add_paragraph('가입 화면이 나오면 차례대로:')
doc.add_paragraph()
doc.add_paragraph('  1) "Enter your email"')
doc.add_paragraph('     이메일 주소를 입력하고 "Continue" 클릭')
doc.add_paragraph()
doc.add_paragraph('  2) "Create a password"')
doc.add_paragraph('     비밀번호를 만드세요 (8자 이상, 숫자 포함)')
doc.add_paragraph('     "Continue" 클릭')
doc.add_paragraph()
doc.add_paragraph('  3) "Enter a username"')
doc.add_paragraph('     사용자 이름을 만드세요 (영어, 숫자, - 만 사용 가능)')
doc.add_paragraph('     예시: my-first-code, student-kim, etf-learner')
doc.add_paragraph('     TIP: 이 이름이 내 GitHub 주소가 돼요! (github.com/내이름)')
doc.add_paragraph('     "Continue" 클릭')
doc.add_paragraph()
doc.add_paragraph('  4) "Would you like to receive product updates?"')
doc.add_paragraph('     이메일 수신 여부 - "n"을 입력해도 괜찮아요')
doc.add_paragraph('     "Continue" 클릭')
doc.add_paragraph()
doc.add_paragraph('  5) 퍼즐 인증이 나오면 화면의 퍼즐을 풀어주세요')
doc.add_paragraph()
doc.add_paragraph('  6) "Create account" 클릭')
doc.add_paragraph()
doc.add_paragraph('  7) 이메일 확인')
doc.add_paragraph('     이메일함에서 GitHub가 보낸 메일을 열어 인증 코드(6자리 숫자)를 확인하세요')
doc.add_paragraph('     그 숫자를 화면에 입력하면 가입 완료!')
doc.add_paragraph()
doc.add_paragraph('가입 완료! 이제 GitHub 회원이에요!')
doc.add_paragraph()

add_heading_styled('Step 3. 새 Repository(저장소) 만들기 (5분)', level=2)
doc.add_paragraph('GitHub에 로그인한 상태에서:')
doc.add_paragraph()
doc.add_paragraph('  1) 오른쪽 위의 "+" 버튼 클릭')
doc.add_paragraph('  2) "New repository" 클릭')
doc.add_paragraph()
doc.add_paragraph('새 저장소 만들기 화면:')
doc.add_paragraph()
doc.add_paragraph('  - "Repository name" (저장소 이름):')
add_code_block('etf-youtube-dashboard')
doc.add_paragraph()
doc.add_paragraph('  - "Description" (설명, 선택사항):')
add_code_block('ETF YouTube Data Dashboard - My First Project')
doc.add_paragraph()
doc.add_paragraph('  - "Public" 선택 (다른 사람도 볼 수 있게)')
doc.add_paragraph('    (비공개로 하고 싶다면 "Private" 선택)')
doc.add_paragraph()
doc.add_paragraph('  - 나머지 옵션은 건드리지 마세요!')
doc.add_paragraph()
doc.add_paragraph('  - 맨 아래 초록색 "Create repository" 버튼 클릭!')
doc.add_paragraph()
doc.add_paragraph('빈 저장소가 만들어졌어요!')
doc.add_paragraph('화면에 명령어들이 보일 텐데, 우리가 직접 따라할 거예요.')
doc.add_paragraph()

add_heading_styled('Step 4. 내 코드를 GitHub에 올리기 (10분)', level=2)
doc.add_paragraph('Git Bash를 열고 프로젝트 폴더로 이동:')
add_code_block('cd ~/etf-dashboard')
doc.add_paragraph()

doc.add_paragraph('아래 명령어를 한 줄씩 차례로 입력하세요:')
doc.add_paragraph()
doc.add_paragraph('1) Git 저장소 시작하기:')
add_code_block('git init')
doc.add_paragraph('>> "이 폴더를 Git으로 관리할게!"라는 뜻')
doc.add_paragraph()

doc.add_paragraph('2) 내 이름과 이메일 등록하기:')
doc.add_paragraph('(GitHub 가입할 때 쓴 이름과 이메일을 넣으세요)')
add_code_block('git config user.name "내이름"\ngit config user.email "내이메일@gmail.com"')
doc.add_paragraph('>> Git에게 "이 코드를 만든 사람은 나야"라고 알려주는 것')
doc.add_paragraph()

doc.add_paragraph('3) .env 파일은 올리지 않도록 설정하기:')
add_code_block('echo ".env" > .gitignore')
doc.add_paragraph('>> .env에는 API 키(비밀번호)가 있으니까 인터넷에 올리면 안 돼요!')
doc.add_paragraph('>> .gitignore 파일은 "이 파일은 올리지 마"라고 Git에게 알려주는 목록')
doc.add_paragraph()

doc.add_paragraph('4) 모든 파일을 올릴 준비하기:')
add_code_block('git add .')
doc.add_paragraph('>> "이 폴더의 모든 파일을 포장해줘"라는 뜻 (마지막에 점(.) 꼭 입력!)')
doc.add_paragraph()

doc.add_paragraph('5) 저장(커밋)하기:')
add_code_block('git commit -m "My first ETF dashboard project"')
doc.add_paragraph('>> "이 상태를 기억해줘"라는 뜻')
doc.add_paragraph('>> 따옴표 안의 메시지는 "이때 뭘 했는지" 메모하는 것')
doc.add_paragraph()

doc.add_paragraph('6) GitHub 저장소와 연결하기:')
doc.add_paragraph('!! "내깃허브이름"을 본인의 GitHub 사용자 이름으로 바꾸세요!')
add_code_block('git remote add origin https://github.com/내깃허브이름/etf-youtube-dashboard.git')
doc.add_paragraph()

doc.add_paragraph('7) GitHub에 올리기!:')
add_code_block('git branch -M main\ngit push -u origin main')
doc.add_paragraph()

doc.add_paragraph('GitHub 로그인 창이 나타날 수 있어요:')
doc.add_paragraph('  - 브라우저가 열리면 "Authorize" 또는 "허용" 클릭')
doc.add_paragraph('  - 또는 GitHub 사용자 이름과 비밀번호 입력')
doc.add_paragraph()
doc.add_paragraph('성공하면 이런 메시지가 나와요:')
add_code_block('To https://github.com/내깃허브이름/etf-youtube-dashboard.git\n * [new branch]      main -> main')
doc.add_paragraph()

add_heading_styled('Step 5. GitHub에서 내 프로젝트 확인하기! (2분)', level=2)
doc.add_paragraph('브라우저에서 아래 주소로 이동하세요 (내깃허브이름을 본인 이름으로):')
add_code_block('https://github.com/내깃허브이름/etf-youtube-dashboard')
doc.add_paragraph()
doc.add_paragraph('내가 만든 파일들이 GitHub 웹사이트에 예쁘게 보여요!')
doc.add_paragraph('이 주소를 친구에게 공유하면, 친구도 내 프로젝트를 볼 수 있습니다!')
doc.add_paragraph()
doc.add_paragraph('TIP: .env 파일은 올라가지 않은 것을 확인하세요!')
doc.add_paragraph('TIP: (비밀번호가 인터넷에 노출되면 안 되니까요)')
doc.add_paragraph()

add_heading_styled('CHECK 체크포인트', level=2)
add_checkpoint([
    'CHECK: GitHub에 가입했나요?',
    'CHECK: etf-youtube-dashboard 저장소를 만들었나요?',
    'CHECK: git push 후 GitHub 웹사이트에서 파일이 보이나요?',
    'CHECK: .env 파일이 GitHub에 올라가지 않았나요? (올라가면 안 돼요!)',
])
doc.add_paragraph()

add_heading_styled('!! 막혔을 때', level=2)
add_error_table([
    ('git push에서 "Authentication failed"', '로그인 실패', '브라우저에서 GitHub 로그인 확인. git push 다시 실행하면 로그인 창이 뜰 수 있어요'),
    ('"remote origin already exists"', '이미 연결됨', 'git remote remove origin 먼저 실행 후 다시 git remote add...'),
    ('"Repository not found"', '저장소 이름 불일치', 'GitHub 웹에서 저장소 이름 확인 후 URL 정확히 입력'),
    ('아무것도 안 올라가요', 'add/commit 미실행', 'git add . -> git commit -m "메시지" -> git push 순서대로 다시'),
    ('"Author identity unknown" 에러', 'Git에 이름/이메일 미등록', 'git config user.name "이름" 과 git config user.email "이메일" 실행'),
])

doc.add_paragraph()
doc.add_paragraph()
add_para('4회차 & 전체 강의 완료! 축하합니다!!!', bold=True, size=16, color=(0, 128, 0))

doc.add_page_break()

# ════════════════════════════════════════════════════
# 전체 로드맵 요약표
# ════════════════════════════════════════════════════
add_heading_styled('전체 로드맵 요약표', level=1)
doc.add_paragraph()
add_para('전체 흐름을 한눈에 확인하세요!', italic=True, size=12)
doc.add_paragraph()

table = doc.add_table(rows=1, cols=5)
table.style = 'Light Grid Accent 1'
table.alignment = WD_TABLE_ALIGNMENT.CENTER
headers = table.rows[0].cells
headers[0].text = '회차'
headers[1].text = '주제'
headers[2].text = '핵심 내용'
headers[3].text = '결과물'
headers[4].text = '사용 도구'

roadmap_data = [
    ('1회차\n(40분)', '터미널부터\nAI 비서까지',
     '- Windows Terminal 설치\n- PowerShell/CMD/Git Bash 비교\n- Git 설치\n- Node.js 설치\n- Claude Code 설치 + 첫 대화!',
     '터미널, Git, Node.js,\nClaude Code 모두 작동\nAI와 첫 대화 성공!',
     'MS Store, Git\nNode.js\nClaude Code'),
    ('2회차\n(35분)', 'Python +\nYouTube API',
     '- Python 설치\n- Google Cloud API 키 발급\n- Claude에게 수집 코드 부탁',
     'Python 작동\nAPI 키 저장\ncollect_data.py 생성',
     'Python\nGoogle Cloud\nClaude Code'),
    ('3회차\n(35분)', '데이터 수집\n+ 대시보드',
     '- 데이터 수집 실행\n- Claude에게 대시보드 부탁\n- 브라우저에서 확인!',
     'etf_youtube_data.json\nindex.html 대시보드\n브라우저에서 확인!',
     'Python\nClaude Code\n브라우저'),
    ('4회차\n(30분)', 'GitHub에\n공유하기',
     '- GitHub 가입\n- Repository 만들기\n- git push로 코드 올리기',
     'GitHub에\n프로젝트 공개!',
     'Git\nGitHub'),
]

for row_data in roadmap_data:
    row = table.add_row().cells
    for i, val in enumerate(row_data):
        row[i].text = val

doc.add_paragraph()
add_para('총 소요 시간: 4회차 = 약 2시간 20분', bold=True, size=13)
doc.add_paragraph()

# 전체 흐름
add_heading_styled('전체 흐름 다이어그램', level=2)
doc.add_paragraph()
add_para(
    '[1회차] 터미널 + Git + Node.js + Claude Code 첫 실행!\n'
    '         |\n'
    '[2회차] Python 설치 + YouTube API 키 + 수집 프로그램 생성\n'
    '         |\n'
    '[3회차] 데이터 수집 실행 + 대시보드 제작 + 브라우저 확인!\n'
    '         |\n'
    '[4회차] GitHub 가입 + 프로젝트 공유 -> 완성!',
    size=11
)

doc.add_paragraph()
doc.add_paragraph()
add_heading_styled('마무리 - 여기까지 완주하신 여러분에게', level=2)
doc.add_paragraph()
doc.add_paragraph('축하합니다! 4회에 걸쳐 여러분은:')
doc.add_paragraph('  1. 터미널(까만 화면)을 두려움 없이 사용하게 되었고')
doc.add_paragraph('  2. PowerShell, CMD, Git Bash의 차이를 알게 되었고')
doc.add_paragraph('  3. Git이라는 개발자 필수 도구를 설치했고')
doc.add_paragraph('  4. Node.js와 Python이라는 프로그래밍 도구를 설치했고')
doc.add_paragraph('  5. AI(Claude Code)와 대화하며 프로그램을 만들었고')
doc.add_paragraph('  6. 유튜브 API로 실제 데이터를 수집했고')
doc.add_paragraph('  7. 그 데이터를 예쁜 대시보드로 시각화했고')
doc.add_paragraph('  8. GitHub에 내 프로젝트를 올려서 세상과 공유했습니다!')
doc.add_paragraph()
doc.add_paragraph('이 경험을 바탕으로 더 다양한 프로젝트에 도전해보세요.')
doc.add_paragraph('Claude Code에게 부탁하면 거의 모든 것을 만들 수 있습니다!')
doc.add_paragraph()
add_para('다음에 도전해볼 것들:', bold=True)
doc.add_paragraph('  - 다른 키워드(예: "비트코인", "부동산")로 데이터 수집해보기')
doc.add_paragraph('  - 대시보드에 검색 기능 추가해보기')
doc.add_paragraph('  - 매일 자동으로 데이터를 모아오게 만들기')
doc.add_paragraph('  - 수집 데이터를 엑셀 파일로 저장하기')
doc.add_paragraph('  - 친구와 함께 GitHub에서 협업해보기')

# ── 저장 ──
desktop = os.path.join(os.path.expanduser('~'), 'Desktop')
filepath = os.path.join(desktop, 'ETF_유튜브_대시보드_강의안.docx')
doc.save(filepath)
print(f'saved: {filepath}')
