import { ComponentFixture, TestBed } from '@angular/core/testing';
import { IsabelaPage } from './isabela.page';

describe('IsabelaPage', () => {
  let component: IsabelaPage;
  let fixture: ComponentFixture<IsabelaPage>;

  beforeEach(() => {
    fixture = TestBed.createComponent(IsabelaPage);
    component = fixture.componentInstance;
    fixture.detectChanges();
  });

  it('should create', () => {
    expect(component).toBeTruthy();
  });
});
