// 데이터는 동일하게 유지
const gitData = [
    { name: 'git init', category: 'basic', desc: '현재 디렉토리를 Git 저장소로 초기화합니다.', example: 'git init' },
    { name: 'git clone', category: 'basic', desc: '원격 저장소를 로컬로 복제합니다.', example: 'git clone https://github.com/user/repo.git' },
    { name: 'git status', category: 'local', desc: '현재 작업 디렉토리의 상태를 확인합니다.', example: 'git status' },
    { name: 'git add', category: 'local', desc: '변경사항을 스테이지(바구니)에 올립니다.', example: 'git add . (모든 파일)\ngit add file.txt (특정 파일)' },
    { name: 'git commit', category: 'local', desc: '스테이지의 변경사항을 로컬 저장소에 저장합니다.', example: 'git commit -m "작업 메시지"' },
    { name: 'git log', category: 'local', desc: '커밋 히스토리를 확인합니다.', example: 'git log --oneline (한 줄로 보기)' },
    { name: 'git push', category: 'remote', desc: '로컬 커밋을 원격 저장소로 보냅니다.', example: 'git push origin main' },
    { name: 'git pull', category: 'remote', desc: '원격 저장소의 최신 내용을 가져와 합칩니다.', example: 'git pull origin main' },
    { name: 'git fetch', category: 'remote', desc: '원격의 바뀐 내용만 확인하고 합치지는 않습니다.', example: 'git fetch origin' },
    { name: 'git branch', category: 'branch', desc: '브랜치 목록을 확인하거나 생성합니다.', example: 'git branch (목록)\ngit branch feature-1 (생성)' },
    { name: 'git checkout', category: 'branch', desc: '브랜치를 전환하거나 파일을 복구합니다.', example: 'git checkout main (전환)\ngit checkout -- file.txt (파일 복구)' },
    { name: 'git merge', category: 'branch', desc: '다른 브랜치를 현재 브랜치로 합칩니다.', example: 'git merge feature-1' },
    { name: 'git stash', category: 'pro', desc: '하던 작업을 임시 보관함에 치워둡니다.', example: 'git stash (보관)\ngit stash pop (꺼내기)' },
    { name: 'git remote', category: 'remote', desc: '원격 저장소 연결 정보를 관리합니다.', example: 'git remote -v (확인)' },
    { name: 'git reset', category: 'pro', desc: '이전 상태로 되돌립니다.', example: 'git reset --hard HEAD~1 (한 단계 전으로)' }
];

document.addEventListener('DOMContentLoaded', function () {
    const contentArea = document.getElementById('content-area');
    const searchInput = document.getElementById('search-input');
    const navItems = document.querySelectorAll('.nav-item');
    const modal = document.getElementById('modal');
    const modalBody = document.getElementById('modal-body');
    const closeBtn = document.querySelector('.close-btn');

    // UI 렌더링 함수
    function renderCards(filter, search) {
        filter = filter || 'all';
        search = (search || '').toLowerCase();

        contentArea.innerHTML = '';

        const filtered = gitData.filter(function (item) {
            const matchesCategory = filter === 'all' || item.category === filter;
            const matchesSearch = item.name.toLowerCase().indexOf(search) !== -1 ||
                item.desc.toLowerCase().indexOf(search) !== -1;
            return matchesCategory && matchesSearch;
        });

        filtered.forEach(function (item) {
            const card = document.createElement('div');
            card.className = 'cmd-card';
            card.innerHTML =
                '<div class="cmd-header">' +
                '<span class="cmd-name">' + item.name + '</span>' +
                '<span class="cmd-tag">' + item.category + '</span>' +
                '</div>' +
                '<p class="cmd-desc">' + item.desc + '</p>' +
                '<div class="cmd-example">' + item.example.replace(/\n/g, '<br>') + '</div>';
            contentArea.appendChild(card);
        });
    }

    // 검색 박스 이벤트
    searchInput.addEventListener('input', function (e) {
        const activeItem = document.querySelector('.nav-item.active');
        const activeCategory = activeItem ? activeItem.getAttribute('data-category') : 'all';
        renderCards(activeCategory, e.target.value);
    });

    // 카테고리 필터 이벤트
    navItems.forEach(function (item) {
        item.addEventListener('click', function () {
            navItems.forEach(function (nb) { nb.classList.remove('active'); });
            item.classList.add('active');
            renderCards(item.getAttribute('data-category'), searchInput.value);
        });
    });

    // 모달 로직
    function showModal(content) {
        modalBody.innerHTML = content;
        modal.style.display = 'block';
        document.body.style.overflow = 'hidden'; // 스크롤 방지
    }

    function hideModal() {
        modal.style.display = 'none';
        document.body.style.overflow = 'auto';
    }

    closeBtn.onclick = hideModal;
    window.onclick = function (e) { if (e.target == modal) hideModal(); };

    // 하단 네비게이션 버튼들
    document.getElementById('btn-home').onclick = function () {
        navItems.forEach(function (nb) { nb.classList.remove('active'); });
        navItems[0].classList.add('active');
        searchInput.value = '';
        renderCards('all', '');
        window.scrollTo(0, 0);
    };

    document.getElementById('btn-flow').onclick = function () {
        const content =
            '<h2 style="margin-bottom:20px; color:var(--primary-light)">기본 작업 흐름</h2>' +
            '<div class="flow-step"><h3>1. git pull</h3><p>팀원의 변경사항을 먼저 가져옵니다.</p></div>' +
            '<div class="flow-step"><h3>2. 코드 작업</h3><p>프로그램을 열심히 수정합니다.</p></div>' +
            '<div class="flow-step"><h3>3. git add .</h3><p>모든 변경사항을 바구니에 담습니다.</p></div>' +
            '<div class="flow-step"><h3>4. git commit</h3><p>작업 내용을 확정하여 저장합니다.</p></div>' +
            '<div class="flow-step"><h3>5. git push</h3><p>서버로 내가 만든 코드를 보냅니다.</p></div>';
        showModal(content);
    };

    document.getElementById('btn-trouble').onclick = function () {
        const content =
            '<h2 style="margin-bottom:20px; color:var(--accent)">자주 발생하는 문제 해결</h2>' +
            '<div style="margin-bottom:15px"><h3 style="color:var(--primary-light)">Q: 실수로 파일을 삭제했어요!</h3><p>A: <code>git checkout -- 파일명</code> 을 입력해 복구하세요.</p></div>' +
            '<div style="margin-bottom:15px"><h3 style="color:var(--primary-light)">Q: 커밋 메시지를 틀렸어요!</h3><p>A: <code>git commit --amend</code> 로 수정 가능합니다.</p></div>' +
            '<div style="margin-bottom:15px"><h3 style="color:var(--primary-light)">Q: 작업 중인데 급하게 브랜치를 바꿔야 해요.</h3><p>A: <code>git stash</code> 로 임시 저장 후 이동하세요.</p></div>';
        showModal(content);
    };

    // 초기 실행
    renderCards();

    // 서비스 워커 등록 (PWA)
    if ('serviceWorker' in navigator) {
        navigator.serviceWorker.register('./sw.js').then(function () {
            console.log('Service Worker Registered');
        });
    }
});
