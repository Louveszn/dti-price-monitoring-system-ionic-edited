import { ComponentFixture, TestBed } from '@angular/core/testing';
import { CagayanPage } from './cagayan.page';

describe('CagayanPage', () => {
  let component: CagayanPage;
  let fixture: ComponentFixture<CagayanPage>;

  beforeEach(() => {
    fixture = TestBed.createComponent(CagayanPage);
    component = fixture.componentInstance;
    fixture.detectChanges();
  });

  it('should create', () => {
    expect(component).toBeTruthy();
  });
});
