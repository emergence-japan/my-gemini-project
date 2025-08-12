document.addEventListener('DOMContentLoaded', function() {
    // Daily cost calculation
    const endDate = new Date('2026-08-12');
    const today = new Date();
    const remainingTime = endDate.getTime() - today.getTime();
    const remainingDays = Math.ceil(remainingTime / (1000 * 60 * 60 * 24));

    if (remainingDays > 0) {
        const dailyCost = Math.ceil(11000 / remainingDays);
        const dailyCostElement1 = document.getElementById('daily-cost-1');
        if (dailyCostElement1) {
            dailyCostElement1.textContent = `1日${dailyCost}円`;
        }
        const dailyCostElement2 = document.getElementById('daily-cost-2');
        if (dailyCostElement2) {
            dailyCostElement2.textContent = `1日${dailyCost}円`;
        }
    } else {
        const dailyCostElement1 = document.getElementById('daily-cost-1');
        if (dailyCostElement1) {
            dailyCostElement1.textContent = '期間終了';
        }
        const dailyCostElement2 = document.getElementById('daily-cost-2');
        if (dailyCostElement2) {
            dailyCostElement2.textContent = '期間終了';
        }
    }

    // FAQ Accordion
    const faqItems = document.querySelectorAll('.faq-item');
    faqItems.forEach(item => {
        const question = item.querySelector('dt');
        question.addEventListener('click', () => {
            const currentlyActive = document.querySelector('.faq-item.active');
            if (currentlyActive && currentlyActive !== item) {
                currentlyActive.classList.remove('active');
            }
            item.classList.toggle('active');
        });
    });

    // Scroll Animation
    const animatedElements = document.querySelectorAll('.card');

    const observer = new IntersectionObserver((entries) => {
        entries.forEach(entry => {
            if (entry.isIntersecting) {
                entry.target.classList.add('is-visible');
                observer.unobserve(entry.target); // Stop observing once visible
            }
        });
    }, {
        threshold: 0.1 // Trigger when 10% of the element is visible
    });

    animatedElements.forEach(el => {
        observer.observe(el);
    });
});