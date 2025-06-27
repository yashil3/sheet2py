function toggleSection(sectionId) {
    const content = document.getElementById(sectionId);
    const header = content.previousElementSibling;
    
    if (content.classList.contains('collapsed')) {
        content.classList.remove('collapsed');
        header.querySelector('.toggle-icon').style.transform = 'rotate(0deg)';
    } else {
        content.classList.add('collapsed');
        header.querySelector('.toggle-icon').style.transform = 'rotate(180deg)';
    }
}

// Initialize collapsed state for formulas sections
document.addEventListener('DOMContentLoaded', function() {
    // Keep vehicle info and customer prefs expanded by default
    // Collapse formulas sections by default
    const formulaSections = document.querySelectorAll('[id^="formulas-"]');
    formulaSections.forEach(section => {
        section.classList.add('collapsed');
        const header = section.previousElementSibling;
        header.querySelector('.toggle-icon').style.transform = 'rotate(180deg)';
    });
});