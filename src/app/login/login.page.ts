import { Component } from '@angular/core';
import { CommonModule } from '@angular/common';
import { FormsModule } from '@angular/forms';
import { Router } from '@angular/router';

import {
  IonContent,
  IonItem,
  IonInput,
  IonButton,
  IonIcon
} from '@ionic/angular/standalone';

import { ToastController } from '@ionic/angular';

@Component({
  selector: 'app-login',
  templateUrl: './login.page.html',
  styleUrls: ['./login.page.scss'],
  standalone: true,
  imports: [
    CommonModule,
    FormsModule,
    IonContent,
    IonItem,
    IonInput,
    IonButton,
    IonIcon
  ]
})
export class LoginPage {
  username = '';
  password = '';
  loading = false;
  showPassword = false;

  private readonly ADMIN_USERNAME = 'admin';
  private readonly ADMIN_PASSWORD = '123456789';

  constructor(
    private toastCtrl: ToastController,
    private router: Router
  ) {}

  togglePassword() {
    this.showPassword = !this.showPassword;
  }

  async login() {
    if (!this.username || !this.password) {
      return this.showToast('Please enter email and password', 'warning');
    }

    this.loading = true;

    setTimeout(async () => {
      if (
        this.username === this.ADMIN_USERNAME &&
        this.password === this.ADMIN_PASSWORD
      ) {
        localStorage.setItem('isLoggedIn', 'true');
        await this.showToast('Welcome Admin!', 'success');
        this.router.navigate(['/menu/home']);
      } else {
        await this.showToast('Invalid credentials', 'danger');
      }
      this.loading = false;
    }, 600);
  }

  private async showToast(message: string, color: string) {
    const toast = await this.toastCtrl.create({
      message,
      duration: 2000,
      color,
      position: 'top'
    });
    toast.present();
  }
}
