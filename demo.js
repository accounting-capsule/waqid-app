/* Fake demo data only. Invented names, no real clients. */
(function (root) {
  function demo() {
    const clients = [
      { name: 'أحمد سمير (تجريبي)', phone: '01000000001', entity: 'مالك فيلا', address: 'الإسكندرية', agreeDate: '2026-01-10', endDate: '', notes: 'بيانات وهمية' },
      { name: 'شركة النور للتطوير (تجريبي)', phone: '01000000002', entity: 'شركة', address: 'القاهرة الجديدة', agreeDate: '2026-02-03', endDate: '', notes: '' },
      { name: 'مها خالد (تجريبي)', phone: '01000000003', entity: 'مالكة وحدة', address: 'طنطا', agreeDate: '2026-03-15', endDate: '', notes: '' }
    ];
    const contracts = [
      { name: 'ترخيص - أحمد سمير (تجريبي)', client: clients[0].name, service: 'ترخيص', date: '2026-01-12', status: 'مكتمل', value: 60000, retention: 0, notes: '' },
      { name: 'إنشاءات - شركة النور (تجريبي)', client: clients[1].name, service: 'إنشاءات', date: '2026-02-05', status: 'قيد المعالجة', value: 250000, retention: 10, notes: 'مرحلتان' },
      { name: 'تشطيبات - مها خالد (تجريبي)', client: clients[2].name, service: 'تشطيبات', date: '2026-03-20', status: 'قيد المعالجة', value: 90000, retention: 5, notes: '' }
    ];
    const bills = [
      { project: contracts[0].name, item: 'دفعة مقدمة', no: '1', date: '2026-01-12', days: 7, value: 30000, notes: '' },
      { project: contracts[0].name, item: 'مستخلص', no: '2', date: '2026-03-01', days: 30, value: 30000, notes: '' },
      { project: contracts[1].name, item: 'دفعة مقدمة', no: '1', date: '2026-02-05', days: 7, value: 75000, notes: '' },
      { project: contracts[1].name, item: 'مستخلص', no: '2', date: '2026-04-10', days: 30, value: 100000, notes: '' },
      { project: contracts[1].name, item: 'مستخلص', no: '3', date: '2026-08-15', days: 30, value: 75000, notes: '' },
      { project: contracts[2].name, item: 'دفعة مقدمة', no: '1', date: '2026-03-20', days: 7, value: 27000, notes: '' },
      { project: contracts[2].name, item: 'مستخلص', no: '2', date: '2026-07-01', days: 30, value: 40000, notes: '' }
    ];
    const journal = [
      { date: '2026-01-13', type: 'إيراد', project: contracts[0].name, item: 'دفعة مقدمة', amount: 30000, method: 'تحويل بنكي', docNo: 'R-001', notes: '' },
      { date: '2026-01-20', type: 'مصروف مشروع', project: contracts[0].name, item: 'رخص وجهات حكومية', amount: 6500, method: 'نقدي', docNo: 'E-101', notes: '' },
      { date: '2026-03-10', type: 'إيراد', project: contracts[0].name, item: 'مستخلص محصل', amount: 30000, method: 'تحويل بنكي', docNo: 'R-002', notes: '' },
      { date: '2026-02-06', type: 'إيراد', project: contracts[1].name, item: 'دفعة مقدمة', amount: 75000, method: 'شيك', docNo: 'R-003', notes: '' },
      { date: '2026-02-20', type: 'مصروف مشروع', project: contracts[1].name, item: 'استشاري فرعي', amount: 12000, method: 'تحويل بنكي', docNo: 'E-102', notes: '' },
      { date: '2026-04-25', type: 'إيراد', project: contracts[1].name, item: 'مستخلص محصل', amount: 60000, method: 'تحويل بنكي', docNo: 'R-004', notes: 'دفعة جزئية من المستخلص 2' },
      { date: '2026-05-02', type: 'مصروف مشروع', project: contracts[1].name, item: 'مقاول باطن', amount: 40000, method: 'شيك', docNo: 'E-103', notes: '' },
      { date: '2026-03-21', type: 'إيراد', project: contracts[2].name, item: 'دفعة مقدمة', amount: 27000, method: 'إنستاباي / محفظة', docNo: 'R-005', notes: '' },
      { date: '2026-04-02', type: 'مصروف مشروع', project: contracts[2].name, item: 'رفع مساحي', amount: 3500, method: 'نقدي', docNo: 'E-104', notes: '' },
      { date: '2026-01-31', type: 'مصروف عمومي', project: '', item: 'إيجار المكتب', amount: 8000, method: 'نقدي', docNo: 'G-001', notes: '' },
      { date: '2026-02-28', type: 'مصروف عمومي', project: '', item: 'مرتبات وأجور', amount: 25000, method: 'تحويل بنكي', docNo: 'G-002', notes: '' },
      { date: '2026-03-31', type: 'مصروف عمومي', project: '', item: 'كهرباء ومياه وإنترنت', amount: 1800, method: 'إنستاباي / محفظة', docNo: 'G-003', notes: '' }
    ];
    return { clients, contracts, bills, journal };
  }
  if (typeof module !== 'undefined') module.exports = demo; else root.WaqidDemo = demo;
})(typeof self !== 'undefined' ? self : this);
