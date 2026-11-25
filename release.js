// Publication and local-preview startup share the same approval checks.
export async function releaseState(brand, terms) {
  const normalized = terms.replace(/\r\n?/g, '\n');
  let digest = '';
  if (globalThis.crypto?.subtle) {
    const bytes = await crypto.subtle.digest('SHA-256', new TextEncoder().encode(normalized));
    digest = Array.from(new Uint8Array(bytes), byte => byte.toString(16).padStart(2, '0')).join('');
  }
  const names = { suite: brand.suite, ...Object.fromEntries(Object.entries(brand.apps).map(([route, app]) => [route, app.name])) };
  const namesApproved = Object.entries(names).every(([key, name]) =>
    typeof name === 'string' && name.trim() && brand.approvedPublicNames?.[key] === name);
  const termsApproved = Boolean(digest && terms.trim() && !/draft placeholder for legal review/i.test(terms)
    && brand.termsVersion && brand.approvedTermsVersion === brand.termsVersion
    && brand.approvedTermsSHA256 === digest);
  return { digest, namesApproved, termsApproved, ready: namesApproved && termsApproved };
}