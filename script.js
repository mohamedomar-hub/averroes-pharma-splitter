function toggleSidebarJS() {
    const sb = document.getElementById('sidebar-custom');
    if (!sb) return;
    sb.classList.toggle('open');
}
function openWhatsAppJS() {
    window.open('https://wa.me/201554694554', '_blank');
}
function navigateTo(anchor) {
    try {
        document.getElementById(anchor).scrollIntoView({behavior: 'smooth', block: 'start'});
    } catch(e) {
        window.location.hash = anchor;
    }
    const sb = document.getElementById('sidebar-custom');
    if (sb && sb.classList.contains('open')) {
        sb.classList.remove('open');
    }
}
function showToastJS(msg) {
    let t = document.getElementById('global-toast');
    if (!t) return;
    t.innerText = msg;
    t.classList.add('show');
    setTimeout(()=>{ t.classList.remove('show'); }, 3000);
}
function backToTopJS() {
    window.scrollTo({top:0, behavior:'smooth'});
}
