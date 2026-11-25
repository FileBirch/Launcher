// Route identifiers and consent storage are independent of display names.
export const brand = {
  suite: 'FileBirch',
  get fullName() {
    const { spreadsheets, documents, presentations } = this.apps;
    return this.suite + ': ' + spreadsheets.name + ', ' + documents.name + ', and ' + presentations.name;
  },
  pageTitle(route) {
    return this.apps[route] ? this.suite + ' ' + this.apps[route].name : this.fullName;
  },
  descriptor: 'Local productivity suite',
  version: '0.1',
  get description() { return this.fullName + '. A browser-based suite for working with local files.'; },
  feedback: 'NaiomiSummersDevelopment@gmail.com',
  stage: 'Early Development',
  warning: 'Errors or missing features are possible in this early release.',
  termsVersion: 'draft-1',
  approvedTermsVersion: '',
  approvedTermsSHA256: '',
  approvedDisclaimers: '',
  // The owner's selected names. This records a product decision, not a legal opinion.
  approvedPublicNames: { suite: 'FileBirch', spreadsheets: 'Tables', documents: 'Papers', presentations: 'Present' },
  art: 'quill-art.js',
  mark: 'quill.svg',
  colors: {
    ink: '#292735', muted: '#605e68', plum: '#523552', deep: '#39283f',
    cream: '#faf8f3', line: '#d8d4ce', copper: '#ab7354'
  },
  apps: {
    spreadsheets: { name: 'Tables', subtitle: 'Spreadsheet editor', formats: '.ods · .xlsx · .csv', available: true },
    documents: { name: 'Papers', subtitle: 'Document editor', formats: '.docx', available: false },
    presentations: { name: 'Present', subtitle: 'Presentation editor', formats: '.pptx', available: false }
  }
};
