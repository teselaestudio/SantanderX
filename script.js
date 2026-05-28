/**
 * Tesela Estudio Presentation Slide Deck Engine
 */

document.addEventListener('DOMContentLoaded', () => {
  // Elements
  const slides = Array.from(document.querySelectorAll('.slide'));
  const slidesWrapper = document.getElementById('slides-wrapper');
  const btnPrev = document.getElementById('btn-prev');
  const btnNext = document.getElementById('btn-next');
  const currentSlideNum = document.getElementById('current-slide-num');
  const totalSlidesNum = document.getElementById('total-slides-num');
  const progressBar = document.getElementById('progress-bar');
  const headerSlideTitle = document.getElementById('header-slide-title');
  
  // Presenter Elements
  const btnPresenter = document.getElementById('btn-presenter');
  const presenterDrawer = document.getElementById('presenter-drawer');
  const btnClosePresenter = document.getElementById('btn-close-presenter');
  const btnPopout = document.getElementById('btn-popout');
  const activeSpeechNotes = document.getElementById('active-speech-notes');
  const currentSlidePreview = document.getElementById('current-slide-preview');
  const nextSlidePreview = document.getElementById('next-slide-preview');
  
  // Fullscreen Elements
  const btnFullscreen = document.getElementById('btn-fullscreen');
  
  // State
  let currentSlideIndex = 0;
  const totalSlides = slides.length;
  let notesWindow = null;

  // Initialize
  totalSlidesNum.textContent = totalSlides;
  updatePresentation();

  // Navigation function
  function goToSlide(index) {
    if (index >= 0 && index < totalSlides) {
      // Remove active class from current
      slides[currentSlideIndex].classList.remove('active');
      
      // Set new index
      currentSlideIndex = index;
      
      // Add active class to new
      slides[currentSlideIndex].classList.add('active');
      
      updatePresentation();
    }
  }

  function nextSlide() {
    if (currentSlideIndex < totalSlides - 1) {
      goToSlide(currentSlideIndex + 1);
    }
  }

  function prevSlide() {
    if (currentSlideIndex > 0) {
      goToSlide(currentSlideIndex - 1);
    }
  }

  // Update presentation view & info
  function updatePresentation() {
    const activeSlide = slides[currentSlideIndex];
    const slideTitle = activeSlide.getAttribute('data-title') || `Diapositiva ${currentSlideIndex + 1}`;
    
    // Update footer slide info
    currentSlideNum.textContent = currentSlideIndex + 1;
    headerSlideTitle.textContent = slideTitle;
    
    // Update progress bar
    const progressPercent = ((currentSlideIndex + 1) / totalSlides) * 100;
    progressBar.style.width = `${progressPercent}%`;
    
    // Update presenter notes
    const notesText = activeSlide.querySelector('.speaker-notes-content').innerHTML;
    activeSpeechNotes.innerHTML = notesText;
    
    // Update presenter previews
    updatePresenterPreviews();
    
    // Broadcast State to LocalStorage (for external presenter window syncing)
    localStorage.setItem('tesela_slide_index', currentSlideIndex);
    
    // Sync external window if open
    if (notesWindow && !notesWindow.closed) {
      notesWindow.postMessage({ type: 'slideChange', index: currentSlideIndex }, '*');
    }
  }

  // Render previews in presenter mode
  function updatePresenterPreviews() {
    const activeSlide = slides[currentSlideIndex];
    const nextSlide = slides[currentSlideIndex + 1];
    
    // Current Slide Preview
    let currentPreviewHtml = createSlidePreviewSummary(activeSlide, currentSlideIndex);
    currentSlidePreview.innerHTML = currentPreviewHtml;
    
    // Next Slide Preview
    if (nextSlide) {
      let nextPreviewHtml = createSlidePreviewSummary(nextSlide, currentSlideIndex + 1);
      nextSlidePreview.innerHTML = nextPreviewHtml;
    } else {
      nextSlidePreview.innerHTML = `
        <div class="preview-empty">
          <p>Fin de la presentación</p>
        </div>
      `;
    }
  }

  // Helper to create slide representation in notes mode
  function createSlidePreviewSummary(slideEl, index) {
    const title = slideEl.getAttribute('data-title');
    
    // Find first image
    const imgEl = slideEl.querySelector('img');
    const imgSrc = imgEl ? imgEl.getAttribute('src') : 'Imagenes/Logo Tesela Estudio.png';
    
    return `
      <div class="slide-preview-summary" style="display: flex; flex-direction: column; height:100%; justify-content:space-between; padding: 12px; font-size: 0.8rem; background: #222; color: #fff;">
        <div style="font-weight:600; color: #10b981; margin-bottom: 5px;">Diapositiva ${index + 1}: ${title}</div>
        <div style="flex-grow: 1; display:flex; align-items:center; justify-content:center; overflow:hidden; border: 1px solid #333; border-radius: 6px; background:#111; margin-bottom: 5px;">
          <img src="${imgSrc}" style="max-width: 100%; max-height: 80px; object-fit: contain; opacity: 0.8;">
        </div>
      </div>
    `;
  }

  // Keyboard navigation
  document.addEventListener('keydown', (e) => {
    switch (e.key) {
      case 'ArrowRight':
      case 'Spacebar':
      case ' ':
      case 'PageDown':
        e.preventDefault();
        nextSlide();
        break;
      case 'ArrowLeft':
      case 'Backspace':
      case 'PageUp':
        e.preventDefault();
        prevSlide();
        break;
      case 'Home':
        e.preventDefault();
        goToSlide(0);
        break;
      case 'End':
        e.preventDefault();
        goToSlide(totalSlides - 1);
        break;
      case 'f':
      case 'F':
        toggleFullscreen();
        break;
      case 'p':
      case 'P':
        togglePresenterDrawer();
        break;
    }
  });

  // Click Event Listeners
  btnPrev.addEventListener('click', prevSlide);
  btnNext.addEventListener('click', nextSlide);
  
  btnPresenter.addEventListener('click', togglePresenterDrawer);
  btnClosePresenter.addEventListener('click', closePresenterDrawer);

  // Presenter Drawer controls
  function togglePresenterDrawer() {
    presenterDrawer.classList.toggle('open');
    btnPresenter.classList.toggle('active');
  }

  function closePresenterDrawer() {
    presenterDrawer.classList.remove('open');
    btnPresenter.classList.remove('active');
  }

  // Popout external window for speech notes
  btnPopout.addEventListener('click', () => {
    const width = 800;
    const height = 600;
    const left = (window.screen.width - width) / 2;
    const top = (window.screen.height - height) / 2;
    
    // Open notes.html
    notesWindow = window.open('notes.html', 'PresenterNotes', 
      `width=${width},height=${height},left=${left},top=${top},menubar=no,toolbar=no,location=no,status=no,resizable=yes,scrollbars=yes`
    );
    
    // Pass slide index to popup once loaded
    if (notesWindow) {
      notesWindow.addEventListener('load', () => {
        notesWindow.postMessage({ type: 'init', index: currentSlideIndex, total: totalSlides }, '*');
      });
    }
  });

  // Handle message communication from popup (in case popup navigates the presentation)
  window.addEventListener('message', (e) => {
    if (e.data && e.data.type === 'navigate') {
      goToSlide(e.data.index);
    }
  });

  // Fullscreen Management
  btnFullscreen.addEventListener('click', toggleFullscreen);

  function toggleFullscreen() {
    if (!document.fullscreenElement) {
      document.documentElement.requestFullscreen()
        .then(() => {
          btnFullscreen.innerHTML = `
            <svg viewBox="0 0 24 24" class="icon"><path fill="currentColor" d="M5 16h3v3h2v-5H5v2zm3-8H5v2h5V5H8v3zm6 11h2v-3h3v-2h-5v5zm2-11V5h-2v5h5V8h-3z"/></svg>
          `;
        })
        .catch(err => {
          console.error(`Error al intentar activar pantalla completa: ${err.message}`);
        });
    } else {
      document.exitFullscreen()
        .then(() => {
          btnFullscreen.innerHTML = `
            <svg viewBox="0 0 24 24" class="icon"><path fill="currentColor" d="M7 14H5v5h5v-2H7v-3zm-2-4h2V7h3V5H5v5zm12 7h-3v2h5v-5h-2v3zM14 5v2h3v3h2V5h-5z"/></svg>
          `;
        });
    }
  }

  // Swipe support for mobile
  let touchStartX = 0;
  let touchEndX = 0;
  
  slidesWrapper.addEventListener('touchstart', (e) => {
    touchStartX = e.changedTouches[0].screenX;
  }, false);
  
  slidesWrapper.addEventListener('touchend', (e) => {
    touchEndX = e.changedTouches[0].screenX;
    handleSwipe();
  }, false);
  
  function handleSwipe() {
    const swipeThreshold = 50;
    if (touchEndX < touchStartX - swipeThreshold) {
      nextSlide(); // Swipe left
    }
    if (touchEndX > touchStartX + swipeThreshold) {
      prevSlide(); // Swipe right
    }
  }
});
