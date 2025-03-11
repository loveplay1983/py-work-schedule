document.getElementById('scheduleForm').addEventListener('submit', function(e) {
    const status = document.getElementById('status');
    status.textContent = 'Generating schedule... Please wait.';
    // Form submission proceeds normally; status updates on download
});