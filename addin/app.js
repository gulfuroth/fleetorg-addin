/* ===== Fleet Org Add-In — app.js ===== */

// =============================================================
// 1. STRINGS — all translatable UI text
// =============================================================
const STRINGS = {
  en: {
    title: 'Fleet Org',
    statusLoading: 'Loading…',
    statusReady: 'Data loaded',
    statusError: 'Error loading data',
    kpiVehicles: 'Vehicles',
    kpiGroups: 'Groups',
    kpiUsers: 'Users',
    kpiLicenses: 'Licenses Used',
    treePanelTitle: 'Group Hierarchy',
    donutPanelTitle: 'Vehicle Status',
    licensePanelTitle: 'License Utilization',
    licenseUsed: 'Used:',
    licenseAvail: 'Available:',
    adminPanelTitle: 'Admin Tools',
    adminCreateTitle: 'Create Group',
    adminGroupNameLabel: 'Group Name',
    adminParentLabel: 'Parent Group',
    adminParentRoot: '— Root —',
    createGroupBtn: 'Create Group',
    adminAssignTitle: 'Bulk Assign Vehicles',
    adminTargetGroupLabel: 'Target Group',
    adminTargetGroupPlaceholder: '— Select group —',
    vehicleSearchPlaceholder: 'Filter vehicles…',
    selectAllVehiclesBtn: 'Select All',
    assignVehiclesBtn: 'Assign Selected',
    adminRemoveTitle: 'Remove from Group',
    adminSourceGroupLabel: 'Source Group',
    adminSourceGroupPlaceholder: '— Select group —',
    loadGroupVehiclesBtn: 'Load Vehicles',
    removeVehiclesBtn: 'Remove Selected',
    loadingText: 'Loading fleet data…',
    statusMoving: 'Moving',
    statusStopped: 'Stopped',
    statusIdling: 'Idling',
    statusDisconnected: 'Disconnected',
    langToggle: 'ES',
    confirmRemove: 'Are you sure you want to remove the selected vehicles from this group? This action cannot be undone.',
    errorNoGroup: 'Please select a group.',
    errorNoName: 'Please enter a group name.',
    errorNoVehicles: 'No vehicles selected.',
    errorNoSourceGroup: 'Please select a source group and load vehicles.',
    msgGroupCreated: 'Group created successfully.',
    msgAssigned: (n) => `${n} vehicle(s) assigned successfully.`,
    msgRemoved: (n) => `${n} vehicle(s) removed successfully.`,
    msgLoadVehicles: (n) => `${n} vehicle(s) in group.`,
  },
  es: {
    title: 'Org. de Flota',
    statusLoading: 'Cargando…',
    statusReady: 'Datos cargados',
    statusError: 'Error al cargar datos',
    kpiVehicles: 'Vehículos',
    kpiGroups: 'Grupos',
    kpiUsers: 'Usuarios',
    kpiLicenses: 'Licencias Usadas',
    treePanelTitle: 'Jerarquía de Grupos',
    donutPanelTitle: 'Estado de Vehículos',
    licensePanelTitle: 'Uso de Licencias',
    licenseUsed: 'Usadas:',
    licenseAvail: 'Disponibles:',
    adminPanelTitle: 'Herramientas Admin',
    adminCreateTitle: 'Crear Grupo',
    adminGroupNameLabel: 'Nombre del Grupo',
    adminParentLabel: 'Grupo Padre',
    adminParentRoot: '— Raíz —',
    createGroupBtn: 'Crear Grupo',
    adminAssignTitle: 'Asignación Masiva',
    adminTargetGroupLabel: 'Grupo Destino',
    adminTargetGroupPlaceholder: '— Seleccione grupo —',
    vehicleSearchPlaceholder: 'Filtrar vehículos…',
    selectAllVehiclesBtn: 'Selec. Todos',
    assignVehiclesBtn: 'Asignar Seleccionados',
    adminRemoveTitle: 'Quitar de Grupo',
    adminSourceGroupLabel: 'Grupo Origen',
    adminSourceGroupPlaceholder: '— Seleccione grupo —',
    loadGroupVehiclesBtn: 'Cargar Vehículos',
    removeVehiclesBtn: 'Quitar Seleccionados',
    loadingText: 'Cargando datos de flota…',
    statusMoving: 'En Marcha',
    statusStopped: 'Detenido',
    statusIdling: 'Ralentí',
    statusDisconnected: 'Desconectado',
    langToggle: 'EN',
    confirmRemove: '¿Está seguro de quitar los vehículos seleccionados de este grupo? Esta acción no se puede deshacer.',
    errorNoGroup: 'Seleccione un grupo.',
    errorNoName: 'Ingrese un nombre para el grupo.',
    errorNoVehicles: 'No hay vehículos seleccionados.',
    errorNoSourceGroup: 'Seleccione un grupo de origen y cargue los vehículos.',
    msgGroupCreated: 'Grupo creado exitosamente.',
    msgAssigned: (n) => `${n} vehículo(s) asignado(s) exitosamente.`,
    msgRemoved: (n) => `${n} vehículo(s) quitado(s) exitosamente.`,
    msgLoadVehicles: (n) => `${n} vehículo(s) en el grupo.`,
  },
};

// =============================================================
// 2. DOM References
// =============================================================
const titleEl             = document.getElementById('title');
const langToggleBtn       = document.getElementById('langToggleBtn');
const statusBadge         = document.getElementById('statusBadge');

const kpiVehiclesVal      = document.getElementById('kpiVehiclesVal');
const kpiVehiclesLabel    = document.getElementById('kpiVehiclesLabel');
const kpiGroupsVal        = document.getElementById('kpiGroupsVal');
const kpiGroupsLabel      = document.getElementById('kpiGroupsLabel');
const kpiUsersVal         = document.getElementById('kpiUsersVal');
const kpiUsersLabel       = document.getElementById('kpiUsersLabel');
const kpiLicensesVal      = document.getElementById('kpiLicensesVal');
const kpiLicensesLabel    = document.getElementById('kpiLicensesLabel');

const treePanelTitle      = document.getElementById('treePanelTitle');
const treeChartEl         = document.getElementById('treeChart');
const donutPanelTitle     = document.getElementById('donutPanelTitle');
const donutChartEl        = document.getElementById('donutChart');

const licensePanelTitle   = document.getElementById('licensePanelTitle');
const licenseBarEl        = document.getElementById('licenseBar');
const licenseUsedLabel    = document.getElementById('licenseUsedLabel');
const licenseUsedVal      = document.getElementById('licenseUsedVal');
const licenseAvailLabel   = document.getElementById('licenseAvailLabel');
const licenseAvailVal     = document.getElementById('licenseAvailVal');
const licensePctVal       = document.getElementById('licensePctVal');

const adminPanel          = document.getElementById('adminPanel');
const adminPanelTitle     = document.getElementById('adminPanelTitle');
const adminCreateTitle    = document.getElementById('adminCreateTitle');
const adminGroupNameLabel = document.getElementById('adminGroupNameLabel');
const adminParentLabel    = document.getElementById('adminParentLabel');
const newGroupName        = document.getElementById('newGroupName');
const parentGroupSelect   = document.getElementById('parentGroupSelect');
const createGroupBtn      = document.getElementById('createGroupBtn');
const createGroupStatus   = document.getElementById('createGroupStatus');

const adminAssignTitle      = document.getElementById('adminAssignTitle');
const adminTargetGroupLabel = document.getElementById('adminTargetGroupLabel');
const assignTargetGroup     = document.getElementById('assignTargetGroup');
const vehicleSearchInput    = document.getElementById('vehicleSearchInput');
const selectAllVehiclesBtn  = document.getElementById('selectAllVehiclesBtn');
const vehicleChecklist      = document.getElementById('vehicleChecklist');
const assignVehiclesBtn     = document.getElementById('assignVehiclesBtn');
const bulkAssignStatus      = document.getElementById('bulkAssignStatus');

const adminRemoveTitle        = document.getElementById('adminRemoveTitle');
const adminSourceGroupLabel   = document.getElementById('adminSourceGroupLabel');
const removeSourceGroup       = document.getElementById('removeSourceGroup');
const loadGroupVehiclesBtn    = document.getElementById('loadGroupVehiclesBtn');
const removeVehicleChecklist  = document.getElementById('removeVehicleChecklist');
const removeVehiclesBtn       = document.getElementById('removeVehiclesBtn');
const removeVehiclesStatus    = document.getElementById('removeVehiclesStatus');

const loadingOverlay = document.getElementById('loadingOverlay');
const loadingText    = document.getElementById('loadingText');

// =============================================================
// 3. Application State
// =============================================================
const state = {
  api: null,
  user: null,
  groups: [],
  devices: [],
  users: [],
  deviceStatuses: [],
  licenses: null,
  groupTree: null,
  lang: 'en',
  adminSelectedVehicles: new Set(),
  adminRemoveSelected: new Set(),
  vehicleFilterText: '',
};

// =============================================================
// 4. Utilities
// =============================================================
function t(key) {
  const s = STRINGS[state.lang] || STRINGS.en;
  const fallback = STRINGS.en;
  return key in s ? s[key] : (key in fallback ? fallback[key] : key);
}

function setStatus(msg, isError = false) {
  statusBadge.textContent = msg;
  statusBadge.className = isError ? 'error' : 'success';
}

function setOpStatus(el, msg, isError = false) {
  el.textContent = msg;
  el.className = 'op-status ' + (isError ? 'error' : 'success');
}

function showLoading() {
  loadingText.textContent = t('loadingText');
  loadingOverlay.removeAttribute('hidden');
}

function hideLoading() {
  loadingOverlay.setAttribute('hidden', '');
}

// =============================================================
// 5. Pure Data Functions
// =============================================================

/**
 * Build a nested group tree annotated with vehicleCount and userCount.
 * Returns a synthetic root node whose children are top-level groups.
 */
function buildGroupTree(groups, devices, users) {
  const byId = new Map();
  for (const g of groups) {
    byId.set(g.id, { ...g, vehicleCount: 0, userCount: 0, children: [] });
  }

  // Count devices per group (devices can belong to multiple groups)
  for (const d of devices) {
    const dGroups = d.groups || [];
    for (const gRef of dGroups) {
      if (byId.has(gRef.id)) {
        byId.get(gRef.id).vehicleCount += 1;
      }
    }
  }

  // Count users per group using companyGroups (preferred) or groups
  for (const u of users) {
    const uGroups = u.companyGroups || u.groups || [];
    for (const gRef of uGroups) {
      if (byId.has(gRef.id)) {
        byId.get(gRef.id).userCount += 1;
      }
    }
  }

  // Wire up parent-child relationships
  const roots = [];
  for (const node of byId.values()) {
    const parentId = node.parent && node.parent.id;
    if (parentId && byId.has(parentId)) {
      byId.get(parentId).children.push(node);
    } else {
      roots.push(node);
    }
  }

  // Propagate counts up the tree so parent tiles reflect descendants
  function propagate(node) {
    for (const child of node.children) {
      propagate(child);
      node.vehicleCount += child.vehicleCount;
      node.userCount    += child.userCount;
    }
  }

  const syntheticRoot = {
    id: '__root__',
    name: 'Fleet',
    vehicleCount: 0,
    userCount: 0,
    children: roots,
  };
  propagate(syntheticRoot);

  return syntheticRoot;
}

/**
 * Parse license data. Falls back to counting active devices if API returns null.
 */
function parseLicenses(licenseResult, devices) {
  if (licenseResult && Array.isArray(licenseResult) && licenseResult.length > 0) {
    let total = 0;
    let used  = 0;
    for (const lic of licenseResult) {
      total += (lic.maxDeviceCount || lic.deviceCount || 0);
      used  += (lic.activeDeviceCount || lic.usedDeviceCount || 0);
    }
    // If parsing produced zeroes, fall back
    if (total > 0) {
      return { used, total, available: Math.max(0, total - used) };
    }
  }
  // Fallback: count active devices as "used", no total known
  const used = devices.filter((d) => d.deviceType !== 'None').length;
  return { used, total: null, available: null };
}

// =============================================================
// 6. Chart instances
// =============================================================
let treeChartInstance    = null;
let donutChartInstance   = null;
let licenseBarInstance   = null;

// =============================================================
// 7. Render Functions
// =============================================================

function renderKpis() {
  kpiVehiclesVal.textContent = state.devices.length.toLocaleString();
  kpiGroupsVal.textContent   = state.groups.length.toLocaleString();
  kpiUsersVal.textContent    = state.users.length.toLocaleString();

  const lic = state.licenses;
  if (lic) {
    kpiLicensesVal.textContent = lic.used.toLocaleString();
  } else {
    kpiLicensesVal.textContent = '—';
  }
}

function toEChartsNodes(nodes) {
  return nodes.map((n) => ({
    name: n.name,
    value: Math.max(n.vehicleCount, 1),
    vehicleCount: n.vehicleCount,
    userCount: n.userCount,
    children: n.children && n.children.length ? toEChartsNodes(n.children) : undefined,
  }));
}

function renderTreeChart() {
  if (!treeChartInstance) {
    treeChartInstance = echarts.init(treeChartEl);
  }

  const root = state.groupTree;
  if (!root) return;

  const data = toEChartsNodes(root.children && root.children.length ? root.children : [root]);

  const tooltipFormatter = (params) => {
    const d = params.data;
    return [
      `<strong>${d.name}</strong>`,
      `${t('kpiVehicles')}: ${d.vehicleCount ?? 0}`,
      `${t('kpiUsers')}: ${d.userCount ?? 0}`,
    ].join('<br>');
  };

  treeChartInstance.setOption({
    tooltip: {
      formatter: tooltipFormatter,
    },
    series: [{
      type: 'treemap',
      data,
      roam: false,
      leafDepth: 1,
      breadcrumb: { show: true, bottom: 8 },
      label: {
        show: true,
        formatter: '{b}',
        fontSize: 12,
        color: '#fff',
      },
      upperLabel: {
        show: true,
        height: 28,
        color: '#fff',
        fontSize: 12,
        fontWeight: 'bold',
      },
      itemStyle: {
        borderColor: '#fff',
        borderWidth: 2,
        gapWidth: 2,
      },
      levels: [
        {
          itemStyle: { borderColor: '#0f766e', borderWidth: 4, gapWidth: 4 },
          upperLabel: { show: false },
        },
        {
          colorSaturation: [0.35, 0.65],
          itemStyle: { borderWidth: 2, gapWidth: 2, borderColorSaturation: 0.8 },
        },
      ],
      emphasis: { focus: 'self' },
    }],
  }, true);
}

function renderDonutChart() {
  if (!donutChartInstance) {
    donutChartInstance = echarts.init(donutChartEl);
  }

  // Classify statuses
  let moving = 0, stopped = 0, idling = 0, disconnected = 0;

  for (const s of state.deviceStatuses) {
    if (!s.isDeviceCommunicating) {
      disconnected += 1;
    } else if (s.isDriving) {
      moving += 1;
    } else if (s.speed > 0) {
      idling += 1;
    } else {
      stopped += 1;
    }
  }

  donutChartInstance.setOption({
    tooltip: {
      trigger: 'item',
      formatter: '{b}: {c} ({d}%)',
    },
    legend: {
      orient: 'vertical',
      right: 8,
      top: 'center',
      textStyle: { fontSize: 12 },
    },
    series: [{
      type: 'pie',
      radius: ['40%', '70%'],
      center: ['40%', '50%'],
      avoidLabelOverlap: false,
      itemStyle: { borderRadius: 4, borderColor: '#fff', borderWidth: 2 },
      label: { show: false },
      emphasis: {
        label: { show: true, fontSize: 13, fontWeight: 'bold' },
      },
      data: [
        { value: moving,       name: t('statusMoving'),       itemStyle: { color: getComputedStyle(document.documentElement).getPropertyValue('--fo-chart-moving').trim()       || '#22c55e' } },
        { value: stopped,      name: t('statusStopped'),      itemStyle: { color: getComputedStyle(document.documentElement).getPropertyValue('--fo-chart-stopped').trim()      || '#f59e0b' } },
        { value: idling,       name: t('statusIdling'),       itemStyle: { color: getComputedStyle(document.documentElement).getPropertyValue('--fo-chart-idling').trim()       || '#3b82f6' } },
        { value: disconnected, name: t('statusDisconnected'), itemStyle: { color: getComputedStyle(document.documentElement).getPropertyValue('--fo-chart-disconnected').trim() || '#94a3b8' } },
      ],
    }],
  }, true);
}

function renderLicenseBar() {
  if (!licenseBarInstance) {
    licenseBarInstance = echarts.init(licenseBarEl);
  }

  const lic = state.licenses;
  if (!lic) {
    licenseBarInstance.setOption({ series: [] }, true);
    licenseUsedVal.textContent  = '—';
    licenseAvailVal.textContent = '—';
    licensePctVal.textContent   = '—';
    return;
  }

  const used  = lic.used;
  const total = lic.total;

  let pct = 0;
  let barColor = '#22c55e';

  if (total && total > 0) {
    pct = Math.round((used / total) * 100);
    if (pct >= 90)      barColor = '#b91c1c';
    else if (pct >= 70) barColor = '#f59e0b';
    else                barColor = '#22c55e';
  }

  licenseUsedVal.textContent  = used.toLocaleString();
  licenseAvailVal.textContent = lic.available !== null ? lic.available.toLocaleString() : '—';
  licensePctVal.textContent   = total ? `${pct}%` : `${used} active`;

  if (!total) {
    // No total known — show a simple informational bar
    licenseBarInstance.setOption({ series: [] }, true);
    return;
  }

  licenseBarInstance.setOption({
    grid: { top: 4, bottom: 4, left: 4, right: 4, containLabel: false },
    xAxis: { type: 'value', max: total, show: false },
    yAxis: { type: 'category', data: [''], show: false },
    series: [{
      type: 'bar',
      data: [used],
      barMaxWidth: '100%',
      showBackground: true,
      backgroundStyle: { color: '#e2e8f0' },
      itemStyle: { color: barColor, borderRadius: [4, 4, 4, 4] },
      label: { show: true, position: 'insideRight', formatter: `${pct}%`, color: '#fff', fontWeight: 'bold' },
    }],
  }, true);
}

// =============================================================
// 8. Data Loading
// =============================================================

function loadAllData() {
  if (!state.api) return;
  showLoading();
  setStatus(t('statusLoading'));

  state.api.multiCall(
    [
      ['Get', { typeName: 'Group' }],
      ['Get', { typeName: 'Device', resultsLimit: 50000 }],
      ['Get', { typeName: 'User',   resultsLimit: 10000 }],
      ['Get', { typeName: 'DeviceStatusInfo', resultsLimit: 50000 }],
      ['Get', { typeName: 'License' }],
    ],
    function(results) {
      hideLoading();
      try {
        state.groups        = results[0] || [];
        state.devices       = results[1] || [];
        state.users         = results[2] || [];
        state.deviceStatuses = results[3] || [];
        const licenseRaw    = results[4];

        state.groupTree = buildGroupTree(state.groups, state.devices, state.users);
        state.licenses  = parseLicenses(licenseRaw, state.devices);

        renderKpis();
        renderTreeChart();
        renderDonutChart();
        renderLicenseBar();

        // Show admin panel for system admins
        if (state.user && state.user.isSystemAdminUser) {
          adminPanel.removeAttribute('hidden');
          populateAdminSelects();
          renderVehicleChecklist(vehicleChecklist, state.devices, state.adminSelectedVehicles);
        }

        setStatus(t('statusReady'));
      } catch (err) {
        setStatus(t('statusError') + ': ' + err.message, true);
      }
    },
    function(err) {
      hideLoading();
      setStatus(t('statusError') + ': ' + (err && err.message ? err.message : String(err)), true);
    }
  );
}

// =============================================================
// 9. Admin Functions
// =============================================================

function populateAdminSelects() {
  const sortedGroups = [...state.groups].sort((a, b) => (a.name || '').localeCompare(b.name || ''));

  function buildOptions(placeholder) {
    return `<option value="">${placeholder}</option>` +
      sortedGroups.map((g) => `<option value="${g.id}">${g.name || g.id}</option>`).join('');
  }

  parentGroupSelect.innerHTML   = `<option value="">${t('adminParentRoot')}</option>` +
    sortedGroups.map((g) => `<option value="${g.id}">${g.name || g.id}</option>`).join('');
  assignTargetGroup.innerHTML   = buildOptions(t('adminTargetGroupPlaceholder'));
  removeSourceGroup.innerHTML   = buildOptions(t('adminSourceGroupPlaceholder'));
}

function renderVehicleChecklist(container, devices, selectedSet) {
  const filterText = state.vehicleFilterText.toLowerCase();
  const filtered   = filterText
    ? devices.filter((d) => (d.name || '').toLowerCase().includes(filterText))
    : devices;

  const sorted = [...filtered].sort((a, b) => (a.name || '').localeCompare(b.name || ''));

  if (sorted.length === 0) {
    container.innerHTML = '<div style="padding:8px 10px;color:var(--fo-muted);font-size:12px;">No vehicles found.</div>';
    return;
  }

  container.innerHTML = sorted.map((d) => {
    const checked = selectedSet.has(d.id) ? 'checked' : '';
    const name    = d.name || d.serialNumber || d.id;
    return `<label><input type="checkbox" data-id="${d.id}" ${checked}> ${name}</label>`;
  }).join('');

  // Sync checkboxes to selectedSet
  container.querySelectorAll('input[type="checkbox"]').forEach((cb) => {
    cb.addEventListener('change', () => {
      const id = cb.dataset.id;
      if (cb.checked) selectedSet.add(id);
      else            selectedSet.delete(id);
    });
  });
}

async function createGroup() {
  const name = newGroupName.value.trim();
  if (!name) {
    setOpStatus(createGroupStatus, t('errorNoName'), true);
    return;
  }

  const parentId = parentGroupSelect.value;
  const entity   = { name };
  if (parentId) {
    entity.parent = { id: parentId };
  }

  createGroupBtn.disabled = true;
  setOpStatus(createGroupStatus, t('statusLoading'));

  state.api.call(
    'Add',
    { typeName: 'Group', entity },
    function() {
      createGroupBtn.disabled = false;
      newGroupName.value = '';
      setOpStatus(createGroupStatus, t('msgGroupCreated'));
      loadAllData(); // refresh all data so the new group appears in the tree
    },
    function(err) {
      createGroupBtn.disabled = false;
      setOpStatus(createGroupStatus, err && err.message ? err.message : String(err), true);
    }
  );
}

/**
 * Execute multiCall batches, max 100 calls per batch.
 * Returns a Promise that resolves when all batches complete (or rejects on first error).
 */
function batchedMultiCall(calls) {
  const CHUNK = 100;
  const chunks = [];
  for (let i = 0; i < calls.length; i += CHUNK) {
    chunks.push(calls.slice(i, i + CHUNK));
  }

  function runChunk(idx) {
    if (idx >= chunks.length) return Promise.resolve();
    return new Promise((resolve, reject) => {
      state.api.multiCall(
        chunks[idx],
        () => resolve(runChunk(idx + 1)),
        (err) => reject(err)
      );
    });
  }

  return runChunk(0);
}

function bulkAssignVehicles() {
  const targetGroupId = assignTargetGroup.value;
  if (!targetGroupId) {
    setOpStatus(bulkAssignStatus, t('errorNoGroup'), true);
    return;
  }
  if (state.adminSelectedVehicles.size === 0) {
    setOpStatus(bulkAssignStatus, t('errorNoVehicles'), true);
    return;
  }

  const selectedIds = Array.from(state.adminSelectedVehicles);
  const deviceMap   = new Map(state.devices.map((d) => [d.id, d]));

  const calls = selectedIds.map((id) => {
    const device = deviceMap.get(id);
    if (!device) return null;

    // Merge: add the target group if not already present
    const existingGroups = (device.groups || []).map((g) => ({ id: g.id }));
    const alreadyIn      = existingGroups.some((g) => g.id === targetGroupId);
    const newGroups      = alreadyIn
      ? existingGroups
      : [...existingGroups, { id: targetGroupId }];

    return ['Set', { typeName: 'Device', entity: { ...device, groups: newGroups } }];
  }).filter(Boolean);

  assignVehiclesBtn.disabled = true;
  setOpStatus(bulkAssignStatus, t('statusLoading'));

  batchedMultiCall(calls).then(() => {
    assignVehiclesBtn.disabled = false;
    const msg = STRINGS[state.lang].msgAssigned(selectedIds.length);
    setOpStatus(bulkAssignStatus, msg);
    state.adminSelectedVehicles.clear();
    loadAllData();
  }).catch((err) => {
    assignVehiclesBtn.disabled = false;
    setOpStatus(bulkAssignStatus, err && err.message ? err.message : String(err), true);
  });
}

function loadGroupVehicles() {
  const sourceGroupId = removeSourceGroup.value;
  if (!sourceGroupId) {
    setOpStatus(removeVehiclesStatus, t('errorNoGroup'), true);
    return;
  }

  // Filter devices that belong to this group
  const groupDevices = state.devices.filter((d) =>
    (d.groups || []).some((g) => g.id === sourceGroupId)
  );

  state.adminRemoveSelected.clear();
  renderVehicleChecklist(removeVehicleChecklist, groupDevices, state.adminRemoveSelected);

  const msg = STRINGS[state.lang].msgLoadVehicles(groupDevices.length);
  setOpStatus(removeVehiclesStatus, msg);
}

function removeVehicles() {
  const sourceGroupId = removeSourceGroup.value;
  if (!sourceGroupId || state.adminRemoveSelected.size === 0) {
    setOpStatus(removeVehiclesStatus, t('errorNoSourceGroup'), true);
    return;
  }

  if (!window.confirm(t('confirmRemove'))) return;

  const selectedIds = Array.from(state.adminRemoveSelected);
  const deviceMap   = new Map(state.devices.map((d) => [d.id, d]));

  const calls = selectedIds.map((id) => {
    const device = deviceMap.get(id);
    if (!device) return null;

    // Filter out the source group from the device's groups
    const newGroups = (device.groups || [])
      .filter((g) => g.id !== sourceGroupId)
      .map((g) => ({ id: g.id }));

    return ['Set', { typeName: 'Device', entity: { ...device, groups: newGroups } }];
  }).filter(Boolean);

  removeVehiclesBtn.disabled = true;
  setOpStatus(removeVehiclesStatus, t('statusLoading'));

  batchedMultiCall(calls).then(() => {
    removeVehiclesBtn.disabled = false;
    const msg = STRINGS[state.lang].msgRemoved(selectedIds.length);
    setOpStatus(removeVehiclesStatus, msg);
    state.adminRemoveSelected.clear();
    removeVehicleChecklist.innerHTML = '';
    loadAllData();
  }).catch((err) => {
    removeVehiclesBtn.disabled = false;
    setOpStatus(removeVehiclesStatus, err && err.message ? err.message : String(err), true);
  });
}

// =============================================================
// 10. i18n — Apply language and update all text nodes
// =============================================================

// Static map of elementId → string key for text that can be set via textContent
const TEXT_MAP = {
  title:                'title',
  langToggleBtn:        'langToggle',
  kpiVehiclesLabel:     'kpiVehicles',
  kpiGroupsLabel:       'kpiGroups',
  kpiUsersLabel:        'kpiUsers',
  kpiLicensesLabel:     'kpiLicenses',
  treePanelTitle:       'treePanelTitle',
  donutPanelTitle:      'donutPanelTitle',
  licensePanelTitle:    'licensePanelTitle',
  licenseUsedLabel:     'licenseUsed',
  licenseAvailLabel:    'licenseAvail',
  adminCreateTitle:     'adminCreateTitle',
  adminGroupNameLabel:  'adminGroupNameLabel',
  adminParentLabel:     'adminParentLabel',
  createGroupBtn:       'createGroupBtn',
  adminAssignTitle:     'adminAssignTitle',
  adminTargetGroupLabel:'adminTargetGroupLabel',
  selectAllVehiclesBtn: 'selectAllVehiclesBtn',
  assignVehiclesBtn:    'assignVehiclesBtn',
  adminRemoveTitle:     'adminRemoveTitle',
  adminSourceGroupLabel:'adminSourceGroupLabel',
  loadGroupVehiclesBtn: 'loadGroupVehiclesBtn',
  removeVehiclesBtn:    'removeVehiclesBtn',
  loadingText:          'loadingText',
};

// Keys that are placeholder attributes
const PLACEHOLDER_MAP = {
  vehicleSearchInput:  'vehicleSearchPlaceholder',
};

function applyLang(lang) {
  state.lang = lang;

  for (const [elId, strKey] of Object.entries(TEXT_MAP)) {
    const el = document.getElementById(elId);
    if (el) {
      // For the title, preserve the inline SVG child
      if (elId === 'title') {
        const svgChild = el.querySelector('svg');
        el.textContent = t(strKey);
        if (svgChild) el.prepend(svgChild);
      } else {
        el.textContent = t(strKey);
      }
    }
  }

  for (const [elId, strKey] of Object.entries(PLACEHOLDER_MAP)) {
    const el = document.getElementById(elId);
    if (el) el.placeholder = t(strKey);
  }

  // Update admin panel title (has SVG child)
  if (adminPanelTitle) {
    const svg = adminPanelTitle.querySelector('svg');
    adminPanelTitle.textContent = t('adminPanelTitle');
    if (svg) adminPanelTitle.prepend(svg);
  }

  // Refresh admin selects option labels
  if (parentGroupSelect && parentGroupSelect.options[0]) {
    parentGroupSelect.options[0].textContent = t('adminParentRoot');
  }
  if (assignTargetGroup && assignTargetGroup.options[0]) {
    assignTargetGroup.options[0].textContent = t('adminTargetGroupPlaceholder');
  }
  if (removeSourceGroup && removeSourceGroup.options[0]) {
    removeSourceGroup.options[0].textContent = t('adminSourceGroupPlaceholder');
  }

  // Re-render charts so labels update
  if (state.groupTree) renderTreeChart();
  if (state.deviceStatuses.length) renderDonutChart();
  if (state.licenses !== undefined) renderLicenseBar();
}

// =============================================================
// 11. Event Listeners
// =============================================================

langToggleBtn.addEventListener('click', () => {
  const next = state.lang === 'en' ? 'es' : 'en';
  applyLang(next);
});

createGroupBtn.addEventListener('click', createGroup);

assignVehiclesBtn.addEventListener('click', bulkAssignVehicles);

selectAllVehiclesBtn.addEventListener('click', () => {
  const filterText = state.vehicleFilterText.toLowerCase();
  const filtered = filterText
    ? state.devices.filter((d) => (d.name || '').toLowerCase().includes(filterText))
    : state.devices;

  const allSelected = filtered.every((d) => state.adminSelectedVehicles.has(d.id));
  if (allSelected) {
    filtered.forEach((d) => state.adminSelectedVehicles.delete(d.id));
  } else {
    filtered.forEach((d) => state.adminSelectedVehicles.add(d.id));
  }
  renderVehicleChecklist(vehicleChecklist, state.devices, state.adminSelectedVehicles);
});

vehicleSearchInput.addEventListener('input', () => {
  state.vehicleFilterText = vehicleSearchInput.value;
  renderVehicleChecklist(vehicleChecklist, state.devices, state.adminSelectedVehicles);
});

loadGroupVehiclesBtn.addEventListener('click', loadGroupVehicles);

removeVehiclesBtn.addEventListener('click', removeVehicles);

// =============================================================
// 12. Resize Handler
// =============================================================

window.addEventListener('resize', () => {
  if (treeChartInstance)  treeChartInstance.resize();
  if (donutChartInstance) donutChartInstance.resize();
  if (licenseBarInstance) licenseBarInstance.resize();
});

// =============================================================
// 13. MyGeotab Lifecycle Export
// =============================================================

geotab.addin.FleetOrg = function () {
  function initialize(freshApi, freshState, callback) {
    state.api  = freshApi;
    state.user = freshState.user;
    applyLang(state.lang);
    loadAllData();
    callback();
  }

  function focus(freshApi, freshState) {
    state.api  = freshApi;
    state.user = freshState.user;
    loadAllData(); // refresh on every focus
  }

  function blur() {
    // Dispose all chart instances to prevent memory leaks
    [treeChartInstance, donutChartInstance, licenseBarInstance].forEach((c) => {
      if (c) c.dispose();
    });
    treeChartInstance  = null;
    donutChartInstance = null;
    licenseBarInstance = null;
  }

  return { initialize, focus, blur };
};
