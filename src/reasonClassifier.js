const s = v => v == null ? '' : String(v).trim()
const n = v => s(v).toLowerCase()

// October 2026 decline taxonomy.
// Priority: explicit reason semantics -> structured gateway text -> decline code fallback.
// The classifier deliberately avoids broad rules such as "issuing bank" or "fraud"
// before more specific AVS / 3DS / CVV / validation checks.
export function reasonInfo(rawReason, rawCode = '') {
  const raw = s(rawReason)
  const text = n(raw)
  const code = n(rawCode)
  if (!raw) return { reason: 'Missing Decline Reason', type: 'Data Quality' }

  // A declined row carrying an approval message is a source-data inconsistency.
  if (text.includes('approved or completed successfully')) {
    return { reason: 'Data Quality / Status Inconsistency', type: 'Data Quality' }
  }

  // Customer / cardholder actions.
  if (/canceled by the user|cancelled by customer|cancelled by user|payment aborted by customer|blocked by cardholder|cardholder\/contact cardholder/.test(text)) {
    return { reason: 'Customer Cancellation / Cardholder Block', type: 'Soft' }
  }

  // 3DS / authentication.
  if (/3d\s*secure|3-d secure|3dsecure|three.?d|threedresult|authenticationvalue|authentication value|unauthenticated|user challenge was unsuccessful|cres\.transstatus|emv 3ds|additional sca|customer authentication required|strong customer authentication|authentication is cancelled|timed out at the acs|cryptogram is mandatory|cryptogram|eci data|sg_threedsbrowser/.test(text)) {
    return { reason: '3DS / Authentication Issue', type: 'Soft' }
  }

  // AVS / billing-address verification.
  if (/avs check|billingpostcode|billingcountryiso2a|billing address|billingcountry|billingdetails|postcode given|billingdetails\.street|billingstreet/.test(text)) {
    return { reason: 'AVS / Billing Verification', type: 'Soft' }
  }

  // CVV / card verification.
  if (/\bcvv\b|\bcvc\b|card verification value|invalid card security code|security code/.test(text)) {
    return { reason: 'CVV / Card Verification', type: 'Hard' }
  }
  if (text.includes('verification data failed')) return { reason: 'Verification Data Failed', type: 'Hard' }

  // Funding / issuer response.
  if (/insufficient fund|not sufficient funds|available funds are insufficient|amount exceeds credit/.test(text)) {
    return { reason: 'Insufficient Fund', type: 'Hard' }
  }
  if (text.includes('do not honor') || text.includes('do not honour')) {
    return { reason: 'Do Not Honor', type: 'Hard' }
  }

  // Fraud and risk.
  if (/suspected fraud|fraud\/security|fraud|antifraud|anti-fraud|fraud screening|suspecting manipulation/.test(text)) {
    return { reason: 'Fraud / Risk Screening', type: 'Soft' }
  }
  if (/risk management|riskmanagement|risk rule failed/.test(text)) {
    return { reason: 'Risk Management Check', type: 'Soft' }
  }

  // Card / BIN restrictions.
  if (/bin blacklisted|blacklisted bin|bin block|banned card|banned|blocked card|restricted card|card is temporarily blocked|new cardholder not activated|blocked, first used|blocked or card not activated|blocked due to a previous chargeback|transaction declined \(account blocked\)|transaction declined \(card blocked\)/.test(text)) {
    return { reason: 'Blocked / Restricted Card', type: 'Soft' }
  }
  if (/lost card|lost\/stolen|stolen card|pick.?up card|pickup card|retrieve the card|card lost/.test(text)) {
    return { reason: 'Lost / Stolen Card', type: 'Hard' }
  }
  if (text === 'expired' || text === 'expired card' || text.includes('expired card') || text.includes('invalid expiration date') || text.includes('networktoken') && text.includes('expired')) {
    return { reason: 'Expired Card', type: 'Hard' }
  }

  // Permission / account / transaction restrictions.
  if (/transaction not permitted|request type is not permitted|not permitted to cardholder|not permitted on card|transaction not allowed at terminal|paymenthandle.*not permitted/.test(text)) {
    return { reason: 'Transaction Not Permitted', type: 'Hard' }
  }
  if (/account closed|closed account|account blocked|account temporarily|invalid account|account not yet activated|account not activated|registration is not confirmed|registration is not valid|account data|account or bank details incorrect|no checking account|no to account|invalid\/nonexistent account|card registration failed/.test(text)) {
    return { reason: 'Account / Registration Issue', type: 'Hard' }
  }

  // Limits / retry controls.
  if (/exceeds amount|exceeds approval amount|exceeds withdrawal|exceeds frequency|daily limit|limit exceeded|decline limit exceeded|card limit reached|pin tries exceeded|allowable number of pin-entry tries|too many invalid tries/.test(text)) {
    return { reason: 'Limit / Frequency Exceeded', type: 'Hard' }
  }
  if (/duplicate merchant reference|duplicate transaction|duplicated transaction/.test(text)) {
    return { reason: 'Duplicate Transaction', type: 'Soft' }
  }
  if (/reattempt not permitted|stop retries|update card or cardholder data before reattempt|re-enter transaction|try again later|retry transaction|bank has requested that you retry/.test(text)) {
    return { reason: 'Retry / Reattempt Issue', type: 'Soft' }
  }

  // Compliance / policy.
  if (/policy|violation of law|compliance_violation|compliance violation|revocation of authorisation|revocation of all authorisations|stop all merchants|stop all future payments|chargeback/.test(text)) {
    return { reason: 'Policy / Compliance', type: 'Soft' }
  }
  if (/security violation|security failure/.test(text)) {
    return { reason: 'Security Violation', type: 'Soft' }
  }

  // Merchant / configuration.
  if (/invalid merchant|merchant account configuration|merchant account is not configured|not configured|merchant system error|invalid configuration data|invalid terminal configuration|wrong acceptor key|terminal configuration/.test(text)) {
    return { reason: 'Merchant / Configuration Issue', type: 'Soft' }
  }
  if (/cc company|credit card company|unrecognized credit card company|not allowed to process this cc company/.test(text)) {
    return { reason: 'Credit Card Company Issue', type: 'Hard' }
  }

  // Input / validation errors.
  if (/validation-error|must not have additional properties|must not be empty|invalid field|incorrect parameters|parameters given are incorrect|found wrong xor missing fields|size must be between|invalid value|illegal character|recipientstate|recipientcity|descriptor must match pattern|mandatory fields are missing|required order parameters are invalid or omitted|card\.holdername|card\.cardexpiry|should not be more than|unsupported characters|bad request|format error|phone number must be between|phone number must|invalid or missing parameter/.test(text)) {
    return { reason: 'Validation / Data Error', type: 'Hard' }
  }

  // System / gateway / PSP errors.
  if (/system malfunction|system error|general error|unexpected system error|communication error|external processing gateway|gateway has rejected|connector\/acquirer|cannot create the charge due to one or more errors|declined due to error|unknown error|error occurred|bank's malfunction|bank malfunction|transaction not found on psp|external error in processing|api call result|content is empty|timeout\/retry|a payment was not performed during this session|session expired|gateway timeout|internal error|content-type application\/json/.test(text)) {
    return { reason: 'System / Gateway / PSP Error', type: 'Soft' }
  }

  // Cryptographic / PIN.
  if (/cryptographic|pin key error|incorrect pin/.test(text)) {
    return { reason: 'Cryptographic / PIN Error', type: 'Hard' }
  }

  // Generic declines.
  if (/generic decline|common decline|declined by payment system/.test(text) || ['deny', 'denied', 'declined'].includes(text)) {
    return { reason: 'Generic / Unspecified Decline', type: 'Hard' }
  }

  // Authorization systems.
  if (/authorization system|authorization/.test(text)) {
    return { reason: 'Authorization System Issue', type: 'Soft' }
  }

  // Invalid transaction / amount.
  if (/invalid transaction|transaction declined \(invalid amount\)/.test(text)) {
    return { reason: 'Invalid Transaction', type: 'Hard' }
  }
  if (/invalid amount/.test(text)) return { reason: 'Invalid Amount', type: 'Hard' }

  // Card lifecycle.
  if (/instrument_declined|life cycle|lifecycle|card number|missing or invalid card type|incorrect card data/.test(text)) {
    return { reason: 'Card / Card Lifecycle Issue', type: 'Hard' }
  }

  // Generic issuer wording.
  if (/issuing bank|issuer|card issuer|issuer\/cardholder|not supported by issuer|blocked by issuer|contact card issuer|refer to card issuer/.test(text)) {
    return { reason: 'Issuing Bank / Card Issuer', type: 'Hard' }
  }

  // Decline-code fallback. Do not use ambiguous codes until reason semantics have been checked.
  const codeMap = {
    '51': ['Insufficient Fund', 'Hard'], '51.02': ['Insufficient Fund', 'Hard'], '51.43': ['Insufficient Fund', 'Hard'],
    '05': ['Do Not Honor', 'Hard'], '05.02': ['Do Not Honor', 'Hard'],
    '59': ['Fraud / Risk Screening', 'Soft'], '83.01': ['Fraud / Risk Screening', 'Soft'], '83.02': ['Fraud / Risk Screening', 'Soft'],
    '61': ['Limit / Frequency Exceeded', 'Hard'], '61.02': ['Limit / Frequency Exceeded', 'Hard'], '65': ['Limit / Frequency Exceeded', 'Hard'],
    '57': ['Transaction Not Permitted', 'Hard'], '5c': ['Transaction Not Permitted', 'Hard'], '58': ['Transaction Not Permitted', 'Hard'],
    '54': ['Expired Card', 'Hard'], '63': ['Security Violation', 'Soft'], '63.01': ['Security Violation', 'Soft'], '89': ['Security Violation', 'Soft'],
    '93': ['Policy / Compliance', 'Soft'], '12': ['Invalid Transaction', 'Hard'], '13': ['Invalid Amount', 'Hard'], '14': ['Account / Registration Issue', 'Hard'],
    '16': ['Fraud / Risk Screening', 'Soft'], '22': ['System / Gateway / PSP Error', 'Soft'], '91': ['Issuer Unavailable', 'Soft'], '96': ['System / Gateway / PSP Error', 'Soft'],
    '79': ['Retry / Reattempt Issue', 'Soft'], '9g': ['Blocked / Restricted Card', 'Soft'], 'n7': ['CVV / Card Verification', 'Hard'], '6p': ['Verification Data Failed', 'Hard'],
    '75': ['Limit / Frequency Exceeded', 'Hard'], '78': ['Account / Registration Issue', 'Hard'], '76': ['Account / Registration Issue', 'Hard'],
    '31': ['Credit Card Company Issue', 'Hard'], '41': ['Lost / Stolen Card', 'Hard'], '43': ['Lost / Stolen Card', 'Hard'],
    '17': ['Customer Cancellation / Cardholder Block', 'Soft'], '30': ['Validation / Data Error', 'Hard'],
    '80': ['System / Gateway / PSP Error', 'Soft'], '81': ['Cryptographic / PIN Error', 'Hard'], '85': ['System / Gateway / PSP Error', 'Soft'],
    '0': ['Generic / Unspecified Decline', 'Hard'], '422': ['Generic / Unspecified Decline', 'Hard']
  }
  if (codeMap[code]) return { reason: codeMap[code][0], type: codeMap[code][1] }

  // Gateway/vendor-specific strings: keep them categorized and visible rather than hiding them as Unclassified.
  if (/gw:|gateway|connector|psp|patra|cptrx|cpsca|cpaut|cprou/.test(text)) {
    return { reason: 'System / Gateway / PSP Error', type: 'Soft' }
  }
  return { reason: 'Other / Vendor-Specific Decline', type: 'Soft' }
}
