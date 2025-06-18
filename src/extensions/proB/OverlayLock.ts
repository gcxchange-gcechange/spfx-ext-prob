export class OverlayLock {
  private overlayId = 'globalOverlayLock';
  private styleId = 'globalOverlayLockStyles';
  private locked = false;
  private checkIntervalId: number | undefined;

  constructor() {
    
  }

  private injectStyles(): void {
    if (!document.getElementById(this.styleId)) {
      const style = document.createElement('style');
      style.id = this.styleId;
      style.innerHTML = `
        #${this.overlayId} {
          position: fixed;
          top: 0;
          left: 0;
          width: 100vw;
          height: 100vh;
          background-color: rgba(0, 0, 0, 0.1);
          z-index: 2147483647;
          pointer-events: all;
          display: none;
        }
      `;
      document.head.appendChild(style);
    }
  }

  private createOverlay(): void {
    if (!document.getElementById(this.overlayId)) {
      const overlay = document.createElement('div');
      overlay.id = this.overlayId;
      document.body.appendChild(overlay);
    }
  }

  public lock(): void {
    this.injectStyles();
    this.createOverlay();

    this.locked = true;
    this.validateLock();

    if (!this.checkIntervalId) {
      this.checkIntervalId = window.setInterval(() => {
        this.validateLock();
      }, 500);
    }
  }

  public unlock(): void {
    this.locked = false;

    const overlay = document.getElementById(this.overlayId);
    if (overlay) {
      overlay.style.display = 'none';
    }

    if (this.checkIntervalId) {
      clearInterval(this.checkIntervalId);
      this.checkIntervalId = undefined;
    }
  }

  public destroy(): void {
    this.locked = false;

    if (this.checkIntervalId) {
      clearInterval(this.checkIntervalId);
      this.checkIntervalId = undefined;
    }

    const overlay = document.getElementById(this.overlayId);
    if (overlay && overlay.parentElement) {
      overlay.parentElement.removeChild(overlay);
    }

    const style = document.getElementById(this.styleId);
    if (style && style.parentElement) {
      style.parentElement.removeChild(style);
    }
  }

  private validateLock(): void {
    if (!this.locked) {
      return;
    }

    let overlay = document.getElementById(this.overlayId);

    if (!overlay) {
      this.createOverlay();
      overlay = document.getElementById(this.overlayId);
    }

    if (overlay && overlay.style.display !== 'block') {
      overlay.style.display = 'block';
    }
  }
}