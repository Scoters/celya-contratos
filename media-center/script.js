document.addEventListener('DOMContentLoaded', () => {
    const cards = document.querySelectorAll('.card');
    
    // Smooth 3D tilt effect for each card
    cards.forEach(card => {
        card.addEventListener('mousemove', e => {
            const rect = card.getBoundingClientRect();
            const x = e.clientX - rect.left;
            const y = e.clientY - rect.top;
            
            const centerX = rect.width / 2;
            const centerY = rect.height / 2;
            
            // Calculate tilt amount
            const tiltX = (y - centerY) / 15;
            const tiltY = (centerX - x) / 15;
            
            // Apply transform
            card.style.transform = `perspective(1000px) rotateX(${tiltX}deg) rotateY(${tiltY}deg) translateY(-8px) scale(1.02)`;
        });
        
        // Reset transform on mouse leave
        card.addEventListener('mouseleave', () => {
            card.style.transform = 'perspective(1000px) rotateX(0) rotateY(0) translateY(0) scale(1)';
            // Add a small transition for the snap back
            card.style.transition = 'transform 0.4s cubic-bezier(0.175, 0.885, 0.32, 1.275)';
        });
        
        // Remove transition on enter to make mouse tracking smooth
        card.addEventListener('mouseenter', () => {
            card.style.transition = 'transform 0.1s ease';
        });
    });
});
