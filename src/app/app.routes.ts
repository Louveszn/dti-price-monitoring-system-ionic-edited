import { Routes } from '@angular/router';

export const routes: Routes = [
  {
    path: '',
    redirectTo: 'login',
    pathMatch: 'full',
  },

  // LOGIN
  {
    path: 'login',
    loadComponent: () =>
      import('./login/login.page').then(m => m.LoginPage),
  },

  // DASHBOARD WITH SIDE MENU
  {
    path: 'menu',
    loadComponent: () =>
      import('./layout/menu/menu.page').then(m => m.MenuPage),
    children: [
      {
        path: 'home',
        loadComponent: () =>
          import('./pages/home/home.page').then(m => m.HomePage),
      },
      {
        path: 'cagayan',
        loadComponent: () =>
          import('./pages/cagayan/cagayan.page').then(m => m.CagayanPage),
      },
      {
        path: 'isabela',
        loadComponent: () =>
          import('./pages/isabela/isabela.page').then(m => m.IsabelaPage),
      },
      {
        path: 'quirino',
        loadComponent: () =>
          import('./pages/quirino/quirino.page').then(m => m.QuirinoPage),
      },
      {
        path: 'nueva',
        loadComponent: () =>
          import('./pages/nueva/nueva.page').then(m => m.NuevaPage),
      },
      {
        path: 'about',
        loadComponent: () =>
          import('./about/about.page').then(m => m.AboutPage),
      },
      {
        path: 'batanes',
        loadComponent: () => import('./pages/batanes/batanes.page').then( m => m.BatanesPage)
      },
      {
        path: '',
        redirectTo: 'home',
        pathMatch: 'full',
      },
    ],
  },
  {
    path: 'generated',
    loadComponent: () => import('./generated/generated.page').then( m => m.GeneratedPage)
  },
  // top-level 'about' removed to avoid duplicate route conflicts
];