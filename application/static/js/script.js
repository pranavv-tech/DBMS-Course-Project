document.addEventListener('DOMContentLoaded', () => {
    const dateInputs = document.querySelectorAll('input[type="date"]');
    dateInputs.forEach((input) => {
        if (!input.value) {
            const now = new Date();
            const formatted = now.toISOString().split('T')[0];
            input.value = formatted;
        }
    });
});
