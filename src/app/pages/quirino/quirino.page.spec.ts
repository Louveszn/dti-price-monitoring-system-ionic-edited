import { ComponentFixture, TestBed } from '@angular/core/testing';
import { QuirinoPage } from './quirino.page';

describe('QuirinoPage', () => {
  let component: QuirinoPage;
  let fixture: ComponentFixture<QuirinoPage>;

  beforeEach(() => {
    fixture = TestBed.createComponent(QuirinoPage);
    component = fixture.componentInstance;
    fixture.detectChanges();
  });

  it('should create', () => {
    expect(component).toBeTruthy();
  });
});
