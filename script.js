// script.js
document.getElementById('contact-form').addEventListener('submit', function (e) {
  e.preventDefault();
  const status = document.getElementById('form-status');
  status.textContent = 'Thank you! Your message has been sent.';
  this.reset();
});
