import { Component } from '@angular/core';
import { IonicModule } from '@ionic/angular';
import { CommonModule } from '@angular/common';

interface TeamMember {
  name: string;
  course: string;
  school: string;
  div: string;
  ojt: string;
  image: string;
}

@Component({
  selector: 'app-about',
  templateUrl: './about.page.html',
  styleUrls: ['./about.page.scss'],
  standalone: true,
  imports: [IonicModule, CommonModule]
})
export class AboutPage {
  showWarning = false;
  warningMessage = '';

  team: TeamMember[] = [
    { 
      name: 'Katelyn B. Bangayan', 
      course: 'Bachelor of Science in Computer Science - 4th year', 
      school: 'Cagayan State University - Carig Campus',
      div: 'Consumer Protection Division',
      ojt: 'DTI R2 - On the Job Trainee',
      image: 'assets/kate.jpg' // Add your image path here
    },
    { 
      name: 'Sherilou S. Tan', 
      course: 'Bachelor of Science in Computer Science - 4th year', 
      school: 'Cagayan State University - Carig Campus', 
      div: 'Consumer Protection Division',
      ojt: 'DTI R2 - On the Job Trainee',
      image: 'assets/sherilou.jpg' // Add your image path here
    }
  ];

  onMemberClick(member: TeamMember) {
    // placeholder for future actions
    console.log('member clicked', member);
  }
}