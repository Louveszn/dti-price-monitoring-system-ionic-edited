import { Component, OnDestroy } from '@angular/core';
import { CommonModule } from '@angular/common';
import { IonicModule } from '@ionic/angular';
import { RouterModule, Router } from '@angular/router';

@Component({
  selector: 'app-menu',
  standalone: true,
  imports: [IonicModule, CommonModule, RouterModule],
  templateUrl: './menu.page.html',
  styleUrls: ['./menu.page.scss'],
})
export class MenuPage implements OnDestroy {

  loggingOut = false;
  private forceNavTimer?: any;

  constructor(private router: Router) {}

  logout(): void {
    if (this.loggingOut) return;
    this.loggingOut = true;

    // 1️⃣ Clear local + session storage
    try {
      ['barangay', 'userRole', 'email'].forEach(k =>
        localStorage.removeItem(k)
      );
      sessionStorage.clear();
    } catch {}

    // 2️⃣ Navigate immediately
    this.router.navigateByUrl('/login', { replaceUrl: true });

    // 3️⃣ Hard fallback if router fails
    this.forceNavTimer = setTimeout(() => {
      if (location.pathname !== '/login') {
        window.location.replace('/login');
      }
    }, 300);
  }

  ngOnDestroy(): void {
    if (this.forceNavTimer) {
      clearTimeout(this.forceNavTimer);
    }
  }
}
