import { ComponentFixture, TestBed } from '@angular/core/testing';
import { BatanesPage } from './batanes.page';

describe('BatanesPage', () => {
  let component: BatanesPage;
  let fixture: ComponentFixture<BatanesPage>;

  beforeEach(() => {
    fixture = TestBed.createComponent(BatanesPage);
    component = fixture.componentInstance;
    fixture.detectChanges();
  });

  it('should create', () => {
    expect(component).toBeTruthy();
  });
});
