/**
 * Yoga with Jessica — Test Configuration
 *
 * Central config for all test suites. Update values here when
 * class times, pages, or infrastructure changes.
 */

module.exports = {
  // ========== ENVIRONMENTS ==========
  env: {
    local: {
      baseUrl: 'http://localhost:8080',
      appsScriptUrl: process.env.TEST_APPS_SCRIPT_URL || 'https://script.google.com/macros/s/AKfycbwgatCsIc5cbyF_myJmBl5eK_cLOeR4Ebe1bs62GD8bp9DAB3yH79ll7zmSGhmfW0cd/exec'
    },
    production: {
      baseUrl: 'https://yogawithjessica.com',
      appsScriptUrl: 'https://script.google.com/macros/s/AKfycbzgNn2zYdI2VBA5JeBu0lQnD-JoJnvmYkDM8sHVx3zdNrTwZlGbvKnLc2uYJg9umovD/exec'
    }
  },

  // Which environment to use for tests (change to 'production' for live checks)
  activeEnv: 'local',

  // ========== CLASS SCHEDULE ==========
  classes: {
    'sunday-online': {
      id: 'sunday-online',
      label: 'Sunday Evening — Online via Google Meet',
      day: 0,
      startHour: 18, startMin: 0,
      endHour: 19, endMin: 15,
      duration: '6:00 PM – 7:15 PM PST',
      type: 'online',
      cutoffMinutesAfterStart: 15,
      tags: ['Open to Everyone']
    },
    'tuesday-ccv': {
      id: 'tuesday-ccv',
      label: 'Tuesday Evening — CCV Clubhouse (In Person)',
      day: 2,
      startHour: 18, startMin: 0,
      endHour: 19, endMin: 15,
      duration: '6:00 PM – 7:15 PM PST',
      type: 'inperson',
      cutoffMinutesAfterStart: 15,
      capacity: 10,
      tags: ['CCV Residents Only', 'In Person']
    },
    'wednesday-restorative': {
      id: 'wednesday-restorative',
      label: 'Wednesday Evening — Restorative Yoga (Online)',
      day: 3,
      startHour: 18, startMin: 0,
      endHour: 19, endMin: 15,
      duration: '6:00 PM – 7:15 PM PST',
      type: 'online',
      cutoffMinutesAfterStart: 15,
      tags: ['Restorative', 'Open to Everyone']
    }
  },

  // ========== PAGES ==========
  pages: [
    { path: '/', file: 'index.html', title: 'Yoga with Jessica' },
    { path: '/about.html', title: 'About Me' },
    { path: '/schedule.html', title: 'Class Schedule' },
    { path: '/signup.html', title: 'Sign Up' },
    { path: '/cancel.html', title: 'Cancel Registration' },
    { path: '/waiver.html', title: 'Liability Waiver' },
    { path: '/donations.html', title: 'Donations' },
    { path: '/props.html', title: 'Props' },
    { path: '/privates.html', title: 'Private' }
  ],

  // ========== IMAGES ==========
  images: [
    'profile.jpg',
    'warriorone.jpeg',
    'warriortwo.jpeg',
    'vrksasana.jpeg',
    'blocks.jpg',
    'straps.jpg',
    'yogablankets.jpg',
    'bolsters.jpg',
    'Iyengar Yoga Chair.jpg'
  ],

  // ========== ANALYTICS ==========
  analyticsId: 'G-NVBH3J0L7M',

  // ========== TEST DATA ==========
  testStudent: {
    firstName: 'Test',
    lastName: 'Student',
    email: 'liu.jessica.x+yogatest@gmail.com'
  },

  testGuest: {
    firstName: 'Test',
    lastName: 'Guest'
  },

  // ========== CAPACITY ==========
  maxCapacity: 10,

  // ========== SHEET COLUMNS ==========
  signupColumns: [
    'Timestamp', 'First Name', 'Last Name', 'Email', 'Class',
    'Class Date', 'Class Type', 'Liability Waiver', 'Guest First Name',
    'Guest Last Name', 'Guest Of', 'Cancel Token', 'Device', 'Browser',
    'City', 'State', 'Zip Code'
  ],

  waitlistColumns: [
    'Timestamp', 'First Name', 'Last Name', 'Email', 'Class',
    'Class Date', 'Class Type', 'Guest First Name', 'Guest Last Name',
    'Status', 'Notified At', 'Device', 'Browser', 'City', 'State', 'Zip Code'
  ],

  emailLogColumns: [
    'Timestamp', 'To', 'Subject', 'Body HTML', 'Cancel Token'
  ],

  // Helper: get active environment config
  getEnv() {
    return this.env[this.activeEnv];
  },

  getBaseUrl() {
    return this.getEnv().baseUrl;
  },

  getAppsScriptUrl() {
    return this.getEnv().appsScriptUrl;
  }
};
