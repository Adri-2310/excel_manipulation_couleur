document.addEventListener('DOMContentLoaded', function () {
    var form = document.getElementById('excel-form');
    if (!form) return;

    form.addEventListener('submit', function () {
        var overlay = document.getElementById('loading-overlay');
        if (overlay) {
            overlay.style.display = 'flex';
        }
    });
});