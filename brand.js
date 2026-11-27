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
  version: '0.2 viewer preview',
  get description() { return this.fullName + '. Read-only viewers for local office files.'; },
  feedback: 'NaiomiSummersDevelopment@gmail.com',
  stage: 'Early Development',
  warning: 'Errors or missing features are possible in this early release.',
  termsVersion: '1.0',
  approvedTermsVersion: '1.0',
  approvedTermsSHA256: 'acef3558467238e853de5e22db8475e9ab3ce499965cf9d18878591e70cdbfa7',
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
    spreadsheets: { name: 'Tables', subtitle: 'Read-only spreadsheet viewer', formats: '.ods · .xlsx · .csv', available: true },
    documents: { name: 'Papers', subtitle: 'Read-only document viewer', formats: '.docx · .odt · .txt', available: true },
    presentations: { name: 'Present', subtitle: 'Read-only presentation viewer', formats: '.pptx · .odp', available: true }
  }
};
