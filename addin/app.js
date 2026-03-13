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
    kpiLicenses: 'Active Devices',
    treePanelTitle: 'Group Hierarchy',
    licensePanelTitle: 'Device Plans',
    licenseUsed: 'Active:',
    licenseAvail: 'Inactive:',
    allVehiclesTitle: 'All Vehicles',
    adminCreateTitle: 'Create Group',
    adminGroupNameLabel: 'Group Name',
    adminParentLabel: 'Parent Group',
    adminParentRoot: '— Root —',
    createGroupBtn: 'Create Group',
    adminAssignTitle: 'Assign to Group',
    adminTargetGroupLabel: 'Check vehicles in the table, then assign them to the selected group.',
    adminTargetGroupPlaceholder: '— Select group —',
    vehicleSearchPlaceholder: 'Search: name=Ford AND plan=Pro OR last<7…',
    selectAllVehiclesBtn: 'Select All',
    assignVehiclesBtn: 'Assign to Group',
    adminRemoveTitle: 'Remove from Group',
    adminSourceGroupLabel: 'Select a group in the hierarchy, check vehicles to remove.',
    adminSourceGroupPlaceholder: '— Select group —',
    loadGroupVehiclesBtn: 'Load Vehicles',
    removeVehiclesBtn: 'Remove Selected',
    loadingText: 'Loading Fleet Data',
    stepGroups: 'Groups',
    stepDevices: 'Devices',
    stepUsers: 'Users',
    stepStatuses: 'Vehicle Status',
    stepLicenses: 'Device Plans',
    statusMoving: 'Moving',
    statusStopped: 'Stopped',
    statusIdling: 'Idling',
    statusDisconnected: 'Disconnected',
    langToggle: 'ES',
    hideEmptyGroupsLabel: 'Hide empty groups',
    showSystemGroupsLabel: 'Show system groups',
    showArchivedDevicesLabel: 'Show archived devices',
    stepOdometer: 'Odometer',
    confirmRemove: 'Click again to confirm removal. This cannot be undone.',
    confirmRemoveBtnLabel: 'Confirm Remove',
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
    kpiLicenses: 'Vehículos Activos',
    treePanelTitle: 'Jerarquía de Grupos',
    licensePanelTitle: 'Planes por Vehículo',
    licenseUsed: 'Activos:',
    licenseAvail: 'Inactivos:',
    allVehiclesTitle: 'Todos los Vehículos',
    adminCreateTitle: 'Crear Grupo',
    adminGroupNameLabel: 'Nombre del Grupo',
    adminParentLabel: 'Grupo Padre',
    adminParentRoot: '— Raíz —',
    createGroupBtn: 'Crear Grupo',
    adminAssignTitle: 'Asignar a Grupo',
    adminTargetGroupLabel: 'Marque vehículos en la tabla y asígnelos al grupo seleccionado.',
    adminTargetGroupPlaceholder: '— Seleccione grupo —',
    vehicleSearchPlaceholder: 'Buscar: name=Ford AND plan=Pro OR last<7…',
    selectAllVehiclesBtn: 'Selec. Todos',
    assignVehiclesBtn: 'Asignar a Grupo',
    adminRemoveTitle: 'Quitar de Grupo',
    adminSourceGroupLabel: 'Seleccione un grupo en la jerarquía y marque los vehículos a quitar.',
    adminSourceGroupPlaceholder: '— Seleccione grupo —',
    loadGroupVehiclesBtn: 'Cargar Vehículos',
    removeVehiclesBtn: 'Quitar Seleccionados',
    loadingText: 'Cargando Datos de Flota',
    stepGroups: 'Grupos',
    stepDevices: 'Vehículos',
    stepUsers: 'Usuarios',
    stepStatuses: 'Estado de Vehículos',
    stepLicenses: 'Planes de Vehículo',
    statusMoving: 'En Marcha',
    statusStopped: 'Detenido',
    statusIdling: 'Ralentí',
    statusDisconnected: 'Desconectado',
    langToggle: 'EN',
    hideEmptyGroupsLabel: 'Ocultar grupos vacíos',
    showSystemGroupsLabel: 'Mostrar grupos del sistema',
    showArchivedDevicesLabel: 'Mostrar dispositivos archivados',
    stepOdometer: 'Odómetro',
    confirmRemove: 'Haga clic de nuevo para confirmar. Esta acción no se puede deshacer.',
    confirmRemoveBtnLabel: 'Confirmar eliminación',
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


const vehicleSearchInput = document.getElementById('vehicleSearchInput');
const groupVehicleTable  = document.getElementById('groupVehicleTable');
const vehiclePanelTitle  = document.getElementById('vehiclePanelTitle');

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
  devicePlans: null,
  devicePlanMap: null,
  isAdmin: false,
  groupTree: null,
  lang: 'en',
  editSelectedDevices: new Set(),
  vehicleFilterText: '',
  selectedGroupId: null,
  editMode: false,
  pendingChanges: [],
  editGroupId: null,
  editGroupName: '',
  editLockedGroupId: null,
  hideEmptyGroups: false,
  showSystemGroups: false,
  showArchivedDevices: false,
  treePanelExpanded: false,
  activeDevices: [],
  allDevices: null,
  odometerMap: new Map(),
  tableSort: { column: null, direction: 'asc' },
  columnFilters: {},
  showRelativeDates: true,
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

const LOAD_STEPS = ['groups', 'devices', 'statuses', 'odometer', 'licenses'];

function stepKey(id) {
  return 'step' + id.charAt(0).toUpperCase() + id.slice(1);
}

function initLoadingSteps() {
  const stepsEl = document.getElementById('loadingSteps');
  if (!stepsEl) return;
  stepsEl.innerHTML = LOAD_STEPS.map((id) =>
    `<div class="loading-step loading-step--pending" id="lstep-${id}">` +
    `<span class="step-icon"></span>` +
    `<span class="step-label">${t(stepKey(id))}</span>` +
    `<span class="step-count"></span>` +
    `</div>`
  ).join('');
}

function markStepLoading(id) {
  const el = document.getElementById('lstep-' + id);
  if (!el) return;
  el.className = 'loading-step loading-step--active';
  el.querySelector('.step-label').textContent = t(stepKey(id));
}

function markStepDone(id, count) {
  const el = document.getElementById('lstep-' + id);
  if (!el) return;
  el.className = 'loading-step loading-step--done';
  const countEl = el.querySelector('.step-count');
  if (countEl) countEl.textContent = count !== undefined ? count.toLocaleString() : '';
}

function markStepError(id) {
  const el = document.getElementById('lstep-' + id);
  if (el) el.className = 'loading-step loading-step--error';
}

function showLoading() {
  loadingText.textContent = t('loadingText');
  initLoadingSteps();
  loadingOverlay.removeAttribute('hidden');
}

function hideLoading() {
  loadingOverlay.setAttribute('hidden', '');
}

function escapeHtml(str) {
  if (str == null || str === '') return '—';
  return String(str).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

function formatConnectDate(dateStr) {
  if (!dateStr) return '—';
  const ms = Date.now() - new Date(dateStr).getTime();
  if (ms < 0) return '—';
  if (state.showRelativeDates) {
    const days = Math.floor(ms / 86400000);
    return days === 0 ? 'Today' : `${days}d ago`;
  }
  return new Date(dateStr).toLocaleDateString(undefined, { year: 'numeric', month: 'short', day: 'numeric' });
}

/** Wraps api.call in a Promise */
function apiCall(method, params) {
  return new Promise((resolve, reject) => {
    state.api.call(method, params, resolve, reject);
  });
}

// =============================================================
// 5. Pure Data Functions
// =============================================================

/**
 * Build a nested group tree rooted at the Geotab Company Group.
 * Each node has:
 *   directCount  – vehicles directly assigned to this group
 *   vehicleCount – directCount + sum of all descendants (used for display totals)
 *   userCount    – same propagation logic for users
 * Returns the Company Group node (id: 'GroupCompanyId').
 * Callers should render root.children (not the root itself).
 */
function buildGroupTree(groups, devices, users, showSystemGroups) {
  // Build node map.
  // IMPORTANT: the Geotab API returns each group with a `children` array of
  // {id} references — NOT a `parent` field. We preserve those refs separately
  // so we can wire the tree, then replace children with actual node objects.
  const byId = new Map();
  for (const g of groups) {
    byId.set(g.id, {
      id: g.id,
      name: g.name,
      color: g.color,
      reference: g.reference,
      _childRefs: (g.children || []).map((c) => c.id || c),
      directCount: 0,
      vehicleCount: 0,
      userCount: 0,
      children: [],
    });
  }

  // Wire the tree: replace _childRefs with actual node objects
  for (const node of byId.values()) {
    for (const childId of node._childRefs) {
      const child = byId.get(childId);
      if (child) node.children.push(child);
    }
  }

  // Count direct device memberships.
  // Device.groups contains the group(s) the device is directly assigned to.
  // If the API also returns ancestor groups for a device, de-duplicate by only
  // crediting the most-specific (deepest) group per branch.
  for (const d of devices) {
    const dGroupIds = new Set((d.groups || []).map((g) => g.id));
    for (const gId of dGroupIds) {
      const node = byId.get(gId);
      if (!node) continue;
      const hasChildInGroups = node.children.some((c) => dGroupIds.has(c.id));
      if (!hasChildInGroups) {
        node.directCount += 1;
        node.vehicleCount += 1;
      }
    }
  }

  // Count users per group (direct membership)
  for (const u of users) {
    for (const gRef of (u.companyGroups || u.groups || [])) {
      if (byId.has(gRef.id)) byId.get(gRef.id).userCount += 1;
    }
  }

  // Geotab system group IDs follow the pattern 'Group<UpperCamelCase>Id'.
  // These are auto-classification groups (asset type, powertrain, driver activity, etc.)
  // and pollute the organisational hierarchy view. Remove them and their subtrees.
  const SYS_GROUP = /^Group[A-Z]/;
  function pruneSystemGroups(node) {
    node.children = node.children.filter((c) => !SYS_GROUP.test(c.id));
    for (const child of node.children) pruneSystemGroups(child);
  }

  // Sort children alphabetically at every level
  function sortChildren(node) {
    node.children.sort((a, b) => (a.name || '').localeCompare(b.name || ''));
    for (const child of node.children) sortChildren(child);
  }

  // Propagate totals upward: parent total = direct + sum(children totals)
  function propagate(node) {
    for (const child of node.children) {
      propagate(child);
      node.vehicleCount += child.vehicleCount;
      node.userCount    += child.userCount;
    }
  }

  // The organisational root in Geotab is always 'GroupCompanyId'
  const companyNode = byId.get('GroupCompanyId');
  if (companyNode) {
    if (!showSystemGroups) pruneSystemGroups(companyNode);
    sortChildren(companyNode);
    propagate(companyNode);
    return companyNode;
  }

  // Fallback: synthetic root using groups that have no parent in our set
  const childIdSet = new Set([...byId.values()].flatMap((n) => n._childRefs));
  const roots = [...byId.values()].filter((n) => !childIdSet.has(n.id));
  const syntheticRoot = {
    id: '__root__', name: 'Fleet',
    directCount: 0, vehicleCount: 0, userCount: 0, children: roots, _childRefs: [],
  };
  sortChildren(syntheticRoot);
  propagate(syntheticRoot);
  return syntheticRoot;
}

// Plan keys that represent inactive/non-billable devices
const PLAN_INACTIVE = new Set(['Suspend', 'Terminate']);

// Geotab DevicePlan enum → human-readable label
const PLAN_DISPLAY = {
  'ProPlus':      'Pro Plus',
  'Pro':          'Pro',
  'Base':         'Go Plan',
  'Intermediate': 'Intermediate',
  'RateBasePlan': 'Rate Base',
  'GoTalk':       'GoTalk',
  'Suspend':      'Suspended',
  'Terminate':    'Terminated',
};

// Colour per plan (active plans = teal/green family, inactive = amber/red)
const PLAN_COLORS = {
  'ProPlus':      '#0f766e',
  'Pro':          '#0d9488',
  'Base':         '#22c55e',
  'Intermediate': '#4ade80',
  'RateBasePlan': '#64748b',
  'GoTalk':       '#6ee7b7',
  'Suspend':      '#f59e0b',
  'Terminate':    '#b91c1c',
};

/**
 * Aggregate DevicePlanBillingInfo records into plan counts.
 * Returns { plans: [{key, label, count, active}], total, active }
 */
function parseDevicePlans(planBillingInfo) {
  const counts = {};
  for (const info of (planBillingInfo || [])) {
    const key = info.devicePlan || 'Unknown';
    counts[key] = (counts[key] || 0) + 1;
  }

  const plans = Object.entries(counts)
    .map(([key, count]) => ({
      key,
      label:  PLAN_DISPLAY[key] || key,
      count,
      active: !PLAN_INACTIVE.has(key),
    }))
    .sort((a, b) => b.count - a.count);

  const total  = plans.reduce((s, p) => s + p.count, 0);
  const active = plans.filter((p) => p.active).reduce((s, p) => s + p.count, 0);

  return { plans, total, active };
}

// =============================================================
// 6. Chart instances
// =============================================================
let treeChartInstance    = null;
let mindmapChartInstance = null;
let treeChartData        = null; // last data array passed to setOption, used for node lookup
let vehicleDisplayCount  = 0;   // count of rows currently shown in the table

// =============================================================
// 7. Render Functions
// =============================================================

function renderTreeView() {
  const container = document.getElementById('mindmapChart');
  if (!container) return;
  const root = state.groupTree;
  if (!root) { container.innerHTML = ''; return; }

  function shouldInclude(node) {
    if (!state.hideEmptyGroups) return true;
    // Include if this node or any descendant has vehicles
    if (node.vehicleCount > 0) return true;
    return node.children && node.children.some(shouldInclude);
  }

  function buildNode(node, depth) {
    if (!shouldInclude(node)) return '';
    const visibleChildren = (node.children || []).filter(shouldInclude);
    const hasChildren = visibleChildren.length > 0;
    const isSelected = node.id === state.selectedGroupId;
    const total  = node.vehicleCount || 0;
    const direct = node.directCount  || 0;
    const countLabel = total > 0
      ? `<span class="htree-count" title="${direct} direct">${total}</span>`
      : '';

    if (hasChildren) {
      const childrenHtml = visibleChildren.map((c) => buildNode(c, depth + 1)).join('');
      return `<details class="htree-node" open>
        <summary class="htree-label${isSelected ? ' htree-label--selected' : ''}" data-gid="${node.id}">
          <span class="htree-name">${escapeHtml(node.name || node.id)}</span>
          ${countLabel}
        </summary>
        <div class="htree-children">${childrenHtml}</div>
      </details>`;
    } else {
      return `<div class="htree-leaf${isSelected ? ' htree-label--selected' : ''}" data-gid="${node.id}">
        <span class="htree-name">${escapeHtml(node.name || node.id)}</span>
        ${countLabel}
      </div>`;
    }
  }

  const topNodes = (root.children || []).filter(shouldInclude);
  container.innerHTML = topNodes.map((c) => buildNode(c, 0)).join('');

  container.querySelectorAll('[data-gid]').forEach((el) => {
    el.addEventListener('click', (e) => {
      // Stop bubbling so parent <summary> nodes don't also fire selectGroup
      e.stopPropagation();
      selectGroup(el.dataset.gid);
      // Update selected highlight without full re-render
      container.querySelectorAll('.htree-label--selected').forEach((s) => s.classList.remove('htree-label--selected'));
      el.classList.add('htree-label--selected');
    });
  });
}

function renderKpis() {
  kpiVehiclesVal.textContent = state.devices.length.toLocaleString();
  kpiGroupsVal.textContent   = state.groups.length.toLocaleString();

  const dp = state.devicePlans;
  kpiLicensesVal.textContent = dp ? dp.active.toLocaleString() : '—';
}

function hasAnyVehicles(node) {
  if (node.vehicleCount > 0) return true;
  return node.children && node.children.some(hasAnyVehicles);
}

// Walk option data + ECharts internal tree in parallel to find the internal node
// reference by groupId. treemapRootToNode requires reference equality, so we
// cannot use data from getOption() (which returns clones).
function findInternalTreeNode(optNodes, internalNodes, groupId) {
  for (let i = 0; i < optNodes.length; i++) {
    const opt    = optNodes[i];
    const intern = internalNodes && internalNodes[i];
    if (!intern) continue;
    if (opt.groupId === groupId) return intern;
    if (opt.children && intern.children) {
      const found = findInternalTreeNode(opt.children, intern.children, groupId);
      if (found) return found;
    }
  }
  return null;
}

function toEChartsNodes(nodes) {
  const filtered = state.hideEmptyGroups ? nodes.filter(hasAnyVehicles) : nodes;
  return filtered.map((n) => {
    const visibleChildren = state.hideEmptyGroups ? (n.children || []).filter(hasAnyVehicles) : (n.children || []);
    const children = visibleChildren.length ? toEChartsNodes(visibleChildren) : undefined;
    return {
      name: n.name,
      groupId: n.id,
      value: Math.max(n.vehicleCount, 1),
      vehicleCount: n.vehicleCount,
      directCount: n.directCount,
      childCount: visibleChildren.length,
      children: children && children.length ? children : undefined,
    };
  });
}

function renderTreeChart() {
  if (!treeChartInstance) {
    treeChartInstance = echarts.init(treeChartEl);
  }

  const root = state.groupTree;
  if (!root) return;

  treeChartData = toEChartsNodes(root.children && root.children.length ? root.children : [root]);
  const data = treeChartData;

  const tooltipFormatter = (params) => {
    const d = params.data;
    const total  = d.vehicleCount ?? 0;
    const direct = d.directCount  ?? 0;
    const sub    = total - direct;
    const childCount = d.childCount ?? 0;
    const lines = [`<strong>${d.name}</strong>`];
    const vehicleLine = sub > 0
      ? `${t('kpiVehicles')}: ${total} (${direct} direct + ${sub} in subgroups)`
      : `${t('kpiVehicles')}: ${total}`;
    lines.push(vehicleLine);
    if (childCount > 0) lines.push(`Subgroups: ${childCount}`);
    return lines.join('<br>');
  };

  treeChartInstance.setOption({
    tooltip: {
      formatter: tooltipFormatter,
    },
    series: [{
      type: 'treemap',
      data,
      roam: 'scale',
      leafDepth: 1,
      breadcrumb: { show: true, bottom: 8 },
      label: {
        show: true,
        formatter: (params) => {
          const d = params.data;
          const total  = d.vehicleCount ?? 0;
          const direct = d.directCount  ?? 0;
          const groups = d.childCount   ?? 0;
          const parts = [d.name];
          if (total > 0) {
            parts.push(direct < total ? `${direct} vehicles (${total} total)` : `${direct} vehicles`);
          }
          if (groups > 0) parts.push(`${groups} subgroups`);
          return parts.join('\n');
        },
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

  treeChartInstance.off('click');
  treeChartInstance.on('click', (params) => {
    if (params.data && params.data.groupId) {
      selectGroup(params.data.groupId);
    }
  });
}

function toggleTreePanelExpand() {
  const panel = document.querySelector('.fo-panel[aria-label="Group hierarchy"]');
  const btn   = document.getElementById('btnExpandTree');
  if (!panel || !btn) return;

  state.treePanelExpanded = !state.treePanelExpanded;
  panel.classList.toggle('fo-panel--fullscreen', state.treePanelExpanded);
  document.body.classList.toggle('fo-has-fullscreen', state.treePanelExpanded);

  // Swap icon: expand ↔ compress
  btn.innerHTML = state.treePanelExpanded
    ? `<svg width="14" height="14" viewBox="0 0 14 14" fill="none" stroke="currentColor" stroke-width="1.8">
        <polyline points="5,1 1,1 1,5"/><polyline points="9,13 13,13 13,9"/>
        <line x1="1" y1="1" x2="6" y2="6"/><line x1="13" y1="13" x2="8" y2="8"/>
       </svg>`
    : `<svg width="14" height="14" viewBox="0 0 14 14" fill="none" stroke="currentColor" stroke-width="1.8">
        <polyline points="9,1 13,1 13,5"/><polyline points="5,13 1,13 1,9"/>
        <line x1="13" y1="1" x2="8" y2="6"/><line x1="1" y1="13" x2="6" y2="8"/>
       </svg>`;

  // Let the layout settle, then resize the ECharts treemap to fill the new space
  setTimeout(() => {
    if (treeChartInstance) treeChartInstance.resize();
  }, 50);
}

function selectGroup(groupId) {
  state.selectedGroupId = groupId;
  const g = state.groups.find((g) => g.id === groupId);

  // In edit mode the device panel is locked — only update the sidebar group selector
  if (state.editMode) {
    if (g) {
      state.editGroupId   = g.id;
      state.editGroupName = g.name || g.id;
      const searchEl = document.getElementById('editGroupSearch');
      if (searchEl) searchEl.value = state.editGroupName;
    }
    return; // device table stays frozen
  }

  const title = g ? (g.name || groupId) : t('allVehiclesTitle');
  vehiclePanelTitle.textContent = title;

  // Sync edit group search to follow tree selection (only when not in edit mode)
  if (g) {
    state.editGroupId   = g.id;
    state.editGroupName = g.name || g.id;
    const searchEl = document.getElementById('editGroupSearch');
    if (searchEl) searchEl.value = state.editGroupName;
  }

  refreshVehicleTable();
}

// Collect a group and all its descendant IDs into a Set
function getGroupIdSet(groupId) {
  const ids = new Set();
  function collectAll(node) {
    ids.add(node.id);
    (node.children || []).forEach(collectAll);
  }
  function find(node) {
    if (node.id === groupId) { collectAll(node); return; }
    (node.children || []).forEach(find);
  }
  if (state.groupTree) find(state.groupTree);
  return ids;
}

function refreshVehicleTable() {
  const groupId = state.editMode ? state.editLockedGroupId : state.selectedGroupId;
  let devices;
  if (groupId) {
    const groupIds = getGroupIdSet(groupId);
    devices = state.devices.filter((d) => (d.groups || []).some((g) => groupIds.has(g.id)));
  } else {
    devices = state.devices;
  }
  const selectedSet = state.editMode ? state.editSelectedDevices : null;
  renderVehicleTable(groupVehicleTable, devices, selectedSet, state.vehicleFilterText);
}


function renderPlanPills() {
  const el = document.getElementById('planPills');
  if (!el) return;
  const dp = state.devicePlans;
  if (!dp || !dp.plans.length) { el.innerHTML = ''; return; }
  el.innerHTML = dp.plans.map((p) => {
    const color = PLAN_COLORS[p.key] || '#94a3b8';
    return `<span class="plan-pill" style="background:${color}" title="${p.label}">${p.label}: <strong>${p.count}</strong></span>`;
  }).join('');
}

// =============================================================
// 8. Data Loading
// =============================================================

async function loadAllData() {
  if (!state.api) return;
  showLoading();
  setStatus(t('statusLoading'));

  try {
    // Step 1: Groups — fast, renders the tree shell immediately
    markStepLoading('groups');
    const groups = await apiCall('Get', { typeName: 'Group' });
    state.groups = groups || [];
    markStepDone('groups', state.groups.length);
    state.groupTree = buildGroupTree(state.groups, state.devices, state.users, state.showSystemGroups);
    renderTreeChart();
    renderKpis();

    // Step 2: Devices — use fromDate=now to get only active devices
    markStepLoading('devices');
    const nowIso = new Date().toISOString();
    const devices = await apiCall('Get', { typeName: 'Device', search: { fromDate: nowIso }, resultsLimit: 50000 });
    state.activeDevices = devices || [];
    state.devices = state.activeDevices;
    markStepDone('devices', state.devices.length);
    state.groupTree = buildGroupTree(state.groups, state.devices, state.users, state.showSystemGroups);
    renderTreeChart();
    renderKpis();

    // Step 3: Device status
    markStepLoading('statuses');
    const statuses = await apiCall('Get', { typeName: 'DeviceStatusInfo', resultsLimit: 50000 });
    state.deviceStatuses = statuses || [];
    markStepDone('statuses', state.deviceStatuses.length);

    // Step 5: Odometer — latest StatusData record per device for DiagnosticOdometerId
    markStepLoading('odometer');
    try {
      const thirtyDaysAgo = new Date(Date.now() - 30 * 24 * 60 * 60 * 1000).toISOString();
      const odomData = await apiCall('Get', {
        typeName: 'StatusData',
        search: { diagnosticSearch: { id: 'DiagnosticOdometerId' }, fromDate: thirtyDaysAgo },
        resultsLimit: 50000,
      });
      // Keep most recent record per device (API returns oldest-first, so iterate in reverse)
      const odomMap = new Map();
      for (const rec of (odomData || [])) {
        const devId = rec.device && rec.device.id;
        if (!devId || rec.data == null) continue;
        const existing = odomMap.get(devId);
        if (!existing || rec.dateTime > existing.dateTime) odomMap.set(devId, rec);
      }
      state.odometerMap = odomMap;
      markStepDone('odometer', odomMap.size);
    } catch (e) {
      state.odometerMap = new Map();
      markStepError('odometer');
    }
    refreshVehicleTable();

    // Step 6: Device plan info — read from devicePlans[] field on each Device object
    markStepLoading('licenses');
    const withPlan = state.devices.filter((d) => d.devicePlans && d.devicePlans.length > 0);
    const synth = withPlan.map((d) => ({ devicePlan: d.devicePlans[0], device: { id: d.id } }));
    state.devicePlanMap = new Map(withPlan.map((d) => [d.id, { devicePlan: d.devicePlans[0] }]));
    state.devicePlans = parseDevicePlans(synth);
    markStepDone('licenses', state.devicePlans.total);
    renderPlanPills();
    renderKpis();

    // Detect admin: fetch the current session user and check security groups.
    // freshState.user is unreliable in MyGeotab add-ins; use getSession() instead.
    try {
      const session = await new Promise((res) => state.api.getSession(res));
      const curUsers = await apiCall('Get', { typeName: 'User', search: { name: session.userName }, resultsLimit: 1 });
      const curUser = curUsers && curUsers[0];
      const ADMIN_SG = new Set(['GroupEverythingSecurityId', 'GroupAdministratorSecurityId', 'GroupSupervisorSecurityId']);
      state.isAdmin = !!(curUser && (
        curUser.isAdministrator || curUser.isSystemAdminUser ||
        (curUser.securityGroups || []).some((g) => ADMIN_SG.has(g.id))
      ));
    } catch (_) {
      state.isAdmin = false;
    }

    // Show vehicle panel for all users
    refreshVehicleTable();

    // Show Edit button for admins only
    const editBtn = document.getElementById('btnEditMode');
    if (editBtn) editBtn.hidden = !state.isAdmin;

    setStatus(t('statusReady'));
  } catch (err) {
    setStatus(t('statusError') + ': ' + (err && err.message ? err.message : String(err)), true);
  } finally {
    hideLoading();
  }
}

// =============================================================
// 9. Admin Functions
// =============================================================

// =============================================================
// 9. Edit Mode
// =============================================================

function refreshVehicleCountBadge() {
  const el = document.getElementById('vehicleCountBadge');
  if (!el) return;
  const sel = state.editMode ? state.editSelectedDevices.size : 0;
  el.textContent = sel > 0
    ? `${vehicleDisplayCount} vehicles (${sel} selected)`
    : `${vehicleDisplayCount} vehicles`;
}

function refreshEditBtn() {
  const btn = document.getElementById('btnEditMode');
  if (!btn) return;
  if (!state.editMode) {
    btn.textContent = 'Edit';
    btn.classList.remove('fo-btn--primary');
    btn.classList.add('fo-btn--ghost');
  } else if (state.pendingChanges.length === 0) {
    btn.textContent = 'Back';
    btn.classList.remove('fo-btn--primary');
    btn.classList.add('fo-btn--ghost');
  } else {
    btn.textContent = 'Save';
    btn.classList.add('fo-btn--primary');
    btn.classList.remove('fo-btn--ghost');
  }
}

function setEditMode(active) {
  state.editMode = active;
  state.editLockedGroupId = active ? state.selectedGroupId : null;
  state.editSelectedDevices.clear();
  refreshEditBtn();
  const panel = document.getElementById('editPanel');
  if (panel) panel.hidden = !active;

  if (!active) {
    // Reset staged panel DOM so it's clean on next open
    const stagedPanel = document.getElementById('editStagedPanel');
    const stagedList  = document.getElementById('editStagedList');
    if (stagedPanel) stagedPanel.hidden = true;
    if (stagedList)  stagedList.innerHTML = '';
    // Clear any op status messages
    ['editAddStatus', 'editCreateStatus'].forEach((id) => {
      const el = document.getElementById(id);
      if (el) { el.textContent = ''; el.className = 'op-status'; }
    });
  } else {
    // Pre-fill group search from current tree selection
    const g = state.groups.find((g) => g.id === (state.editGroupId || state.selectedGroupId));
    if (g) {
      state.editGroupId   = g.id;
      state.editGroupName = g.name || g.id;
      const searchEl = document.getElementById('editGroupSearch');
      if (searchEl) searchEl.value = state.editGroupName;
    }
    // Render any already-staged changes (e.g. re-entering after Cancel on modal)
    renderPendingSummary();
  }

  if (!active) {
    // Restore panel title to match current tree selection (may differ from locked group)
    const g = state.groups.find((g) => g.id === state.selectedGroupId);
    vehiclePanelTitle.textContent = g ? (g.name || state.selectedGroupId) : t('allVehiclesTitle');
  }

  refreshVehicleTable();
}

function renderPendingSummary() {
  const changes = state.pendingChanges;
  const panel   = document.getElementById('editStagedPanel');
  const list    = document.getElementById('editStagedList');
  refreshEditBtn();
  if (!panel || !list) return;

  if (!changes.length) { panel.hidden = true; return; }
  panel.hidden = false;

  list.innerHTML = changes.map((c, i) => {
    let text;
    if (c.type === 'add') {
      text = `Add <strong>${c.deviceIds.size}</strong> device(s) → <strong>${escapeHtml(c.groupName)}</strong>`;
    } else if (c.type === 'remove') {
      text = `Remove <strong>${escapeHtml(c.deviceName)}</strong> from <strong>${escapeHtml(c.groupName)}</strong>`;
    } else {
      text = `Create group <strong>${escapeHtml(c.name)}</strong> (parent: <strong>${escapeHtml(c.parentName)}</strong>)`;
    }
    return `<div class="edit-pending-item">
      <span>${text}</span>
      <button class="edit-pending-del" data-idx="${i}" title="Remove">✕</button>
    </div>`;
  }).join('');

  list.querySelectorAll('.edit-pending-del').forEach((btn) => {
    btn.addEventListener('click', () => {
      state.pendingChanges.splice(parseInt(btn.dataset.idx), 1);
      renderPendingSummary();
      refreshVehicleTable();
    });
  });
}

function openConfirmModal() {
  const modal   = document.getElementById('confirmModal');
  const summary = document.getElementById('confirmSummary');
  if (!modal || !summary) return;

  const lines = state.pendingChanges.map((c) => {
    if (c.type === 'add') {
      return `<li>Add <strong>${c.deviceIds.size}</strong> device(s) to group <strong>${escapeHtml(c.groupName)}</strong></li>`;
    } else if (c.type === 'remove') {
      return `<li>Remove <strong>${escapeHtml(c.deviceName)}</strong> from group <strong>${escapeHtml(c.groupName)}</strong></li>`;
    } else {
      return `<li>Create group <strong>${escapeHtml(c.name)}</strong> (parent: <strong>${escapeHtml(c.parentName)}</strong>)</li>`;
    }
  });
  summary.innerHTML = `<ul class="confirm-list">${lines.join('')}</ul>`;
  modal.hidden = false;
}

async function applyPendingChanges() {
  const changes  = [...state.pendingChanges];
  const applyBtn = document.getElementById('confirmApply');
  const modal    = document.getElementById('confirmModal');
  const summary  = document.getElementById('confirmSummary');

  if (applyBtn) { applyBtn.disabled = true; applyBtn.textContent = 'Applying…'; }

  // Helper: create one group via direct callback API (same pattern as original working code)
  function createOneGroup(name, parentId) {
    return new Promise((resolve, reject) => {
      state.api.call(
        'Add',
        { typeName: 'Group', entity: { name, parent: { id: parentId } } },
        resolve,
        (err) => reject(new Error(err && err.message ? err.message : String(err)))
      );
    });
  }

  function showApplyError(msg) {
    if (applyBtn) { applyBtn.disabled = false; applyBtn.textContent = 'Apply Changes'; }
    if (summary) {
      // Remove previous error if any
      summary.querySelectorAll('.apply-error').forEach((el) => el.remove());
      summary.insertAdjacentHTML('beforeend',
        `<p class="op-status error apply-error" style="margin-top:8px">${escapeHtml(msg)}</p>`);
    }
  }

  try {
    const deviceMap = new Map(state.devices.map((d) => [d.id, d]));
    const setCalls  = [];
    const tempIdMap = new Map(); // tempId → real Geotab group ID

    // Process creates first (serial)
    const createChanges = changes.filter((c) => c.type === 'createGroup');
    for (const c of createChanges) {
      await createOneGroup(c.name, c.parentId || 'GroupCompanyId');
    }

    // Re-fetch groups to get real IDs (callback may not return the ID in all SDK versions)
    if (createChanges.length) {
      const freshGroups = await apiCall('Get', { typeName: 'Group' });
      for (const c of createChanges) {
        if (!c.tempId) continue;
        const match = freshGroups.find(
          (g) => g.name === c.name && (g.parent?.id === c.parentId || (!c.parentId && g.parent?.id === 'GroupCompanyId'))
        );
        if (match) tempIdMap.set(c.tempId, match.id);
      }
    }

    // Build Set calls for add/remove, resolving any temp group IDs
    for (const c of changes) {
      if (c.type === 'add') {
        const groupId = tempIdMap.get(c.groupId) || c.groupId;
        if (!groupId || groupId.startsWith('_pending_')) continue; // unresolved temp
        for (const deviceId of c.deviceIds) {
          const device = deviceMap.get(deviceId);
          if (!device) continue;
          const existing = (device.groups || []).map((g) => ({ id: g.id }));
          if (!existing.some((g) => g.id === groupId)) {
            setCalls.push(['Set', { typeName: 'Device', entity: { id: device.id, groups: [...existing, { id: groupId }] } }]);
          }
        }
      } else if (c.type === 'remove') {
        const device = deviceMap.get(c.deviceId);
        if (!device) continue;
        const newGroups = (device.groups || []).filter((g) => g.id !== c.groupId).map((g) => ({ id: g.id }));
        setCalls.push(['Set', { typeName: 'Device', entity: { id: device.id, groups: newGroups } }]);
      }
    }

    if (setCalls.length) await batchedMultiCall(setCalls);

    state.pendingChanges = [];
    if (modal) modal.hidden = true;
    setEditMode(false);
    loadAllData();
  } catch (err) {
    showApplyError(err && err.message ? err.message : String(err));
  }
}

function getPlanLabel(device) {
  if (state.devicePlanMap) {
    const info = state.devicePlanMap.get(device.id);
    if (info && info.devicePlan) return PLAN_DISPLAY[info.devicePlan] || info.devicePlan;
  }
  if (device.devicePlans && device.devicePlans.length > 0) {
    return PLAN_DISPLAY[device.devicePlans[0]] || device.devicePlans[0];
  }
  return '—';
}

// Column definitions for the vehicle table
const VT_COLS = [
  { id: 'name',         label: () => t('kpiVehicles'), sortable: true },
  { id: 'plate',        label: () => 'Plate',          sortable: true },
  { id: 'odometer',     label: () => 'Odometer (km)',  sortable: true, numeric: true },
  { id: 'lastConnect',  label: () => 'Last Connect',   sortable: true, numeric: true },
  { id: 'firstConnect', label: () => 'First Connect',  sortable: true, numeric: true },
  { id: 'plan',         label: () => 'Plan',           sortable: true, filterable: true },
  { id: 'vin',          label: () => 'VIN',            sortable: false },
  { id: 'serial',       label: () => 'Serial No.',     sortable: false },
  { id: 'comment',      label: () => 'Comment',        sortable: false },
];

function getDeviceRow(d, statusMap) {
  const st = statusMap.get(d.id);
  const odomRec = state.odometerMap && state.odometerMap.get(d.id);
  let odoKm = null;
  if (odomRec && odomRec.data != null && odomRec.data > 0) odoKm = Math.round(odomRec.data / 1000);
  else if (st && st.odometer != null && st.odometer > 0) odoKm = Math.round(st.odometer / 1000);

  const lastConnectDate = st ? st.dateTime : null;
  const lastConnectMs   = lastConnectDate ? Date.now() - new Date(lastConnectDate).getTime() : Infinity;
  const firstConnectDate = d.activeFrom || null;
  const firstConnectMs   = firstConnectDate ? Date.now() - new Date(firstConnectDate).getTime() : Infinity;

  return {
    device: d,
    name: (d.name || d.id).toLowerCase(),
    plate: (d.licensePlate || '').toLowerCase(),
    odometer: odoKm,
    lastConnect: lastConnectMs,
    firstConnect: firstConnectMs,
    plan: getPlanLabel(d),
    vin: d.vehicleIdentificationNumber || '',
    serial: d.serialNumber || '',
    comment: d.comment || '',
    lastConnectDate,
    firstConnectDate,
    // display values
    odoDisplay: odoKm != null ? odoKm.toLocaleString() : '—',
    lastConnectDisplay: formatConnectDate(lastConnectDate),
    firstConnectDisplay: formatConnectDate(firstConnectDate),
  };
}

// =============================================================
// SEARCH QUERY PARSER
// Syntax: [col][op][value] AND/OR [col][op][value] ...
// ops: = ~ != < > <= >=
// date cols: value in days; odometer: km; text: substring
// plain text (no col): searches name, plate, vin, serial, comment
// =============================================================
const SEARCH_COL_ALIASES = {
  name: 'name', vehicle: 'name', unit: 'name',
  plate: 'plate', license: 'plate',
  plan: 'plan',
  odo: 'odometer', odometer: 'odometer', km: 'odometer',
  last: 'lastConnect', lastconnect: 'lastConnect',
  first: 'firstConnect', firstconnect: 'firstConnect',
  vin: 'vin', serial: 'serial', comment: 'comment',
};
const SEARCH_TEXT_COLS = new Set(['name', 'plate', 'plan', 'vin', 'serial', 'comment']);
const SEARCH_NUM_COLS  = new Set(['odometer', 'lastConnect', 'firstConnect']);

function parseSearchQuery(text) {
  if (!text || !text.trim()) return [];
  // Split on AND / OR tokens while keeping the operator
  const tokens = text.trim().split(/\s+(AND|OR)\s+/i);
  const clauses = [];
  let pendingLogic = 'AND';
  for (const token of tokens) {
    const upper = token.toUpperCase();
    if (upper === 'AND' || upper === 'OR') { pendingLogic = upper; continue; }
    const m = token.match(/^(\w+)\s*(<=|>=|!=|=|~|<|>)\s*(.+)$/i);
    if (m) {
      const col = SEARCH_COL_ALIASES[m[1].toLowerCase()];
      if (col) {
        clauses.push({ logic: pendingLogic, col, op: m[2], value: m[3].trim() });
        pendingLogic = 'AND';
        continue;
      }
    }
    // No recognised column → plain text across text fields
    clauses.push({ logic: pendingLogic, col: null, op: '~', value: token.trim() });
    pendingLogic = 'AND';
  }
  return clauses;
}

function evalClause(row, { col, op, value }) {
  if (!col) {
    const v = value.toLowerCase();
    return ['name', 'plate', 'vin', 'serial', 'comment'].some((c) =>
      (row[c] || '').toLowerCase().includes(v));
  }
  if (SEARCH_TEXT_COLS.has(col)) {
    const rv = (row[col] || '').toLowerCase();
    const v  = value.toLowerCase();
    if (op === '=' || op === '~') return rv.includes(v);
    if (op === '!=')              return !rv.includes(v);
    return true;
  }
  if (SEARCH_NUM_COLS.has(col)) {
    const num = parseFloat(value);
    if (!isFinite(num)) return true;
    if (col === 'lastConnect' || col === 'firstConnect') {
      if (!isFinite(row[col])) return false; // never connected
      const days = row[col] / 86400000;
      if (op === '=' || op === '~') return Math.abs(days - num) < 1;
      if (op === '!=') return Math.abs(days - num) >= 1;
      if (op === '<')  return days < num;
      if (op === '>')  return days > num;
      if (op === '<=') return days <= num;
      if (op === '>=') return days >= num;
    } else {
      const rv = row[col];
      if (rv == null) return false;
      if (op === '=' || op === '~') return rv === num;
      if (op === '!=') return rv !== num;
      if (op === '<')  return rv < num;
      if (op === '>')  return rv > num;
      if (op === '<=') return rv <= num;
      if (op === '>=') return rv >= num;
    }
  }
  return true;
}

function applySearchFilter(rows, clauses) {
  if (!clauses.length) return rows;
  return rows.filter((row) => {
    let result = evalClause(row, clauses[0]);
    for (let i = 1; i < clauses.length; i++) {
      const clauseResult = evalClause(row, clauses[i]);
      result = clauses[i].logic === 'OR' ? result || clauseResult : result && clauseResult;
    }
    return result;
  });
}

// Returns column config, initialising from VT_COLS defaults if not yet set
function getColumnConfig() {
  if (!state.columnConfig) {
    state.columnConfig = VT_COLS.map((c) => ({ id: c.id, visible: true }));
  }
  return state.columnConfig;
}

// Render a single <td> for a given column id
function renderCell(colId, r, d, groups, planCell, groupMap) {
  switch (colId) {
    case 'name': {
      const namePart = `<div class="vt-name">${escapeHtml(d.name || d.id)}</div>`;
      let subPart = '';
      if (state.editMode && groupMap) {
        const tags = (d.groups || []).map((g) => {
          const gname = groupMap.get(g.id) || g.id;
          const isPending = state.pendingChanges.some(
            (c) => c.type === 'remove' && c.deviceId === d.id && c.groupId === g.id
          );
          return `<span class="vt-group-badge${isPending ? ' vt-group-badge--remove' : ''}"
            data-device-id="${escapeHtml(d.id)}"
            data-device-name="${escapeHtml(d.name || d.id)}"
            data-group-id="${escapeHtml(g.id)}"
            data-group-name="${escapeHtml(gname)}"
            title="${isPending ? 'Click to undo' : 'Click to stage removal'}">${escapeHtml(gname)}</span>`;
        }).join('');
        if (tags) subPart = `<div class="vt-group-badges">${tags}</div>`;
      } else if (groups) {
        subPart = `<div class="vt-sub">${escapeHtml(groups)}</div>`;
      }
      return `<td>${namePart}${subPart}</td>`;
    }
    case 'plate':
      return `<td>${escapeHtml(d.licensePlate || '')}</td>`;
    case 'odometer':
      return `<td class="vt-num">${r.odoDisplay}</td>`;
    case 'lastConnect':
      return `<td class="vt-num vt-filterable" data-filter-col="lastConnect" data-filter-ms="${r.lastConnect}" data-filter-display="${escapeHtml(r.lastConnectDisplay)}">${r.lastConnectDisplay}</td>`;
    case 'firstConnect':
      return `<td class="vt-num vt-filterable" data-filter-col="firstConnect" data-filter-ms="${r.firstConnect}" data-filter-display="${escapeHtml(r.firstConnectDisplay)}">${r.firstConnectDisplay}</td>`;
    case 'plan':
      return planCell;
    case 'vin':
      return `<td class="vt-mono">${escapeHtml(d.vehicleIdentificationNumber || '')}</td>`;
    case 'serial':
      return `<td class="vt-mono">${escapeHtml(d.serialNumber || '')}</td>`;
    case 'comment':
      return `<td>${escapeHtml(d.comment || '')}</td>`;
    default:
      return '<td></td>';
  }
}

function openColumnConfigPanel(anchorBtn) {
  document.getElementById('colConfigPanel')?.remove();
  const cfg = getColumnConfig();

  const items = cfg.map((c, i) => {
    const col = VT_COLS.find((x) => x.id === c.id);
    return `<li class="col-config-item" draggable="true" data-idx="${i}" data-col="${c.id}">
      <span class="col-drag-handle" title="Drag to reorder">⠿</span>
      <input type="checkbox" id="colchk_${c.id}"${c.visible ? ' checked' : ''} data-col="${c.id}">
      <label for="colchk_${c.id}">${col ? col.label() : c.id}</label>
    </li>`;
  }).join('');

  const panel = document.createElement('div');
  panel.id = 'colConfigPanel';
  panel.className = 'col-config-panel';
  panel.innerHTML = `
    <div class="col-config-header">
      <span>Columns</span>
      <button class="col-config-close" title="Close">✕</button>
    </div>
    <ul class="col-config-list">${items}</ul>
    <div class="col-config-footer">
      <button class="fo-btn fo-btn--ghost fo-btn--sm" id="colConfigReset">Reset</button>
    </div>`;

  document.body.appendChild(panel);

  // Position below the anchor, right-aligned
  const rect = anchorBtn.getBoundingClientRect();
  panel.style.top   = (rect.bottom + window.scrollY + 4) + 'px';
  panel.style.right = (window.innerWidth - rect.right + window.scrollX) + 'px';

  // Drag-and-drop reorder
  let dragIdx = null;
  panel.querySelectorAll('.col-config-item').forEach((item) => {
    item.addEventListener('dragstart', (e) => {
      dragIdx = parseInt(item.dataset.idx);
      e.dataTransfer.effectAllowed = 'move';
      item.classList.add('col-config-item--dragging');
    });
    item.addEventListener('dragend', () => item.classList.remove('col-config-item--dragging'));
    item.addEventListener('dragover', (e) => { e.preventDefault(); e.dataTransfer.dropEffect = 'move'; });
    item.addEventListener('drop', (e) => {
      e.preventDefault();
      const dropIdx = parseInt(item.dataset.idx);
      if (dragIdx === null || dragIdx === dropIdx) return;
      const newCfg = [...state.columnConfig];
      const [moved] = newCfg.splice(dragIdx, 1);
      newCfg.splice(dropIdx, 0, moved);
      state.columnConfig = newCfg;
      refreshVehicleTable();
      openColumnConfigPanel(anchorBtn); // re-render panel with new order
    });
  });

  // Checkbox visibility toggle
  panel.querySelectorAll('input[type="checkbox"][data-col]').forEach((chk) => {
    chk.addEventListener('change', () => {
      const entry = state.columnConfig.find((c) => c.id === chk.dataset.col);
      if (entry) { entry.visible = chk.checked; refreshVehicleTable(); }
    });
  });

  // Reset to defaults
  panel.querySelector('#colConfigReset').addEventListener('click', () => {
    state.columnConfig = VT_COLS.map((c) => ({ id: c.id, visible: true }));
    refreshVehicleTable();
    openColumnConfigPanel(anchorBtn);
  });

  // Close button
  panel.querySelector('.col-config-close').addEventListener('click', () => panel.remove());

  // Close on outside click (defer one tick so this click doesn't immediately close)
  setTimeout(() => {
    const handler = (e) => {
      if (!panel.contains(e.target) && e.target !== anchorBtn) {
        panel.remove();
        document.removeEventListener('click', handler);
      }
    };
    document.addEventListener('click', handler);
  }, 0);
}

function renderVehicleTable(container, devices, selectedSet, filterText) {
  const statusMap = new Map(state.deviceStatuses.map((s) => [s.device && s.device.id, s]));

  // Build enriched rows, apply search query and column filters
  let rows = devices.map((d) => getDeviceRow(d, statusMap));
  rows = applySearchFilter(rows, parseSearchQuery(filterText));
  if (state.columnFilters.plan) rows = rows.filter((r) => r.plan === state.columnFilters.plan);
  for (const col of ['lastConnect', 'firstConnect']) {
    const f = state.columnFilters[col];
    if (f) rows = rows.filter((r) => f.mode === 'lte' ? r[col] <= f.ms : r[col] >= f.ms);
  }

  // Sort
  const { column, direction } = state.tableSort;
  if (column) {
    const col = VT_COLS.find((c) => c.id === column);
    const mul = direction === 'asc' ? 1 : -1;
    rows.sort((a, b) => {
      const av = a[column], bv = b[column];
      if (av == null && bv == null) return 0;
      if (av == null) return mul;
      if (bv == null) return -mul;
      if (col && col.numeric) return (av - bv) * mul;
      return String(av).localeCompare(String(bv)) * mul;
    });
  } else {
    rows.sort((a, b) => a.name.localeCompare(b.name));
  }

  vehicleDisplayCount = rows.length;
  refreshVehicleCountBadge();

  if (rows.length === 0) {
    container.innerHTML = '<div class="fo-empty">No vehicles found.</div>';
    return;
  }

  const groupMap = new Map(state.groups.map((g) => [g.id, g.name || g.id]));
  const showCb   = selectedSet != null;

  // Build header with sort indicators
  const sortIcon = (id) => {
    if (state.tableSort.column !== id) return '<span class="vt-sort-icon">⇅</span>';
    return state.tableSort.direction === 'asc'
      ? '<span class="vt-sort-icon vt-sort-icon--active">▲</span>'
      : '<span class="vt-sort-icon vt-sort-icon--active">▼</span>';
  };
  const filterBadge = (colId) => {
    const f = state.columnFilters[colId];
    if (!f) return '';
    const sym = f.mode === 'lte' ? '≤' : '≥';
    return ` <span class="vt-filter-active" data-clear-filter="${colId}" title="Clear filter">${sym}${f.display} ✕</span>`;
  };

  const visibleCols = getColumnConfig()
    .filter((c) => c.visible)
    .map((c) => VT_COLS.find((x) => x.id === c.id))
    .filter(Boolean);

  const allChecked = showCb && rows.length > 0 && rows.every((r) => selectedSet.has(r.device.id));
  const cbHead = showCb
    ? `<th class="vt-cb-col"><input type="checkbox" id="vtSelectAll"${allChecked ? ' checked' : ''} title="Select all"></th>`
    : '';
  const thCells = visibleCols.map((col) => {
    const sort = col.sortable ? ` data-sort="${col.id}"` : '';
    const cls  = col.sortable ? ' class="vt-th-sortable"' : '';
    let extra = '';
    if (col.id === 'plan') extra = filterBadge('plan');
    if (col.id === 'lastConnect')  extra = filterBadge('lastConnect');
    if (col.id === 'firstConnect') extra = filterBadge('firstConnect');
    return `<th${cls}${sort}>${col.label()}${col.sortable ? sortIcon(col.id) : ''}${extra}</th>`;
  }).join('');
  const thead = `<thead><tr>${cbHead}${thCells}</tr></thead>`;

  // Build rows
  const tbodyRows = rows.map((r) => {
    const d = r.device;
    const groups = (d.groups || []).map((g) => groupMap.get(g.id)).filter(Boolean).join(', ');
    const cbCell = showCb
      ? `<td class="vt-cb-col"><input type="checkbox" data-id="${escapeHtml(d.id)}"${selectedSet.has(d.id) ? ' checked' : ''}></td>`
      : '';
    const planCell = state.columnFilters.plan === r.plan
      ? `<td><span class="vt-plan-filtered">${escapeHtml(r.plan)}</span></td>`
      : `<td class="vt-filterable" data-filter-col="plan" data-filter-val="${escapeHtml(r.plan)}">${escapeHtml(r.plan)}</td>`;
    const tdCells = visibleCols.map((col) => renderCell(col.id, r, d, groups, planCell, groupMap)).join('');
    return `<tr>${cbCell}${tdCells}</tr>`;
  }).join('');

  container.innerHTML = `<table class="fo-vehicle-table">${thead}<tbody>${tbodyRows}</tbody></table>`;

  // Sort click handlers
  container.querySelectorAll('[data-sort]').forEach((th) => {
    th.addEventListener('click', (e) => {
      if (e.target.closest('[data-clear-filter]')) return; // handled below
      const col = th.dataset.sort;
      if (state.tableSort.column === col) {
        state.tableSort.direction = state.tableSort.direction === 'asc' ? 'desc' : 'asc';
      } else {
        state.tableSort = { column: col, direction: 'asc' };
      }
      refreshVehicleTable();
    });
  });

  // Column filter: clicking a cell value sets a filter
  container.querySelectorAll('[data-filter-col]').forEach((td) => {
    td.addEventListener('click', () => {
      const col = td.dataset.filterCol;
      if (col === 'plan') {
        const val = td.dataset.filterVal;
        state.columnFilters.plan = state.columnFilters.plan === val ? null : val;
      } else if (col === 'lastConnect' || col === 'firstConnect') {
        const ms = parseFloat(td.dataset.filterMs);
        if (!isFinite(ms)) return;
        const display = td.dataset.filterDisplay;
        const existing = state.columnFilters[col];
        // Toggle off if clicking the same value
        if (existing && existing.ms === ms) { state.columnFilters[col] = null; }
        else {
          // Mode: if this column is sorted ascending (newest first = small ms), user clicked
          // something near the bottom → show ≥ (older); if sorted descending → show ≤ (newer).
          // If not sorted by this col, default to ≤ (show this recent and newer).
          const isSortedAsc  = state.tableSort.column === col && state.tableSort.direction === 'asc';
          const mode = isSortedAsc ? 'gte' : 'lte';
          state.columnFilters[col] = { ms, mode, display };
        }
      }
      refreshVehicleTable();
    });
  });

  // Clear filter ✕ button
  container.querySelectorAll('[data-clear-filter]').forEach((span) => {
    span.addEventListener('click', (e) => {
      e.stopPropagation();
      state.columnFilters[span.dataset.clearFilter] = null;
      refreshVehicleTable();
    });
  });

  // Checkbox handlers (row checkboxes + select-all header)
  if (showCb) {
    const selectAllChk = container.querySelector('#vtSelectAll');
    if (selectAllChk) {
      selectAllChk.addEventListener('change', () => {
        if (selectAllChk.checked) rows.forEach((r) => selectedSet.add(r.device.id));
        else                       rows.forEach((r) => selectedSet.delete(r.device.id));
        refreshVehicleTable();
      });
    }
    container.querySelectorAll('input[type="checkbox"][data-id]').forEach((cb) => {
      cb.addEventListener('change', () => {
        if (cb.checked) selectedSet.add(cb.dataset.id);
        else            selectedSet.delete(cb.dataset.id);
        refreshVehicleCountBadge();
      });
    });
  }

  // Group badge clicks (edit mode removal staging)
  if (state.editMode) {
    container.querySelectorAll('.vt-group-badge').forEach((badge) => {
      // Single-click: toggle removal for this device only
      badge.addEventListener('click', (e) => {
        e.stopPropagation();
        const { deviceId, deviceName, groupId, groupName } = badge.dataset;
        const idx = state.pendingChanges.findIndex(
          (c) => c.type === 'remove' && c.deviceId === deviceId && c.groupId === groupId
        );
        if (idx >= 0) state.pendingChanges.splice(idx, 1);
        else state.pendingChanges.push({ type: 'remove', deviceId, deviceName, groupId, groupName });
        renderPendingSummary();
        refreshVehicleTable();
      });

      // Double-click: bulk-toggle removal for ALL listed devices that have this group
      badge.addEventListener('dblclick', (e) => {
        e.stopPropagation();
        const { groupId, groupName } = badge.dataset;
        // Collect all badges for this group currently in the table
        const allBadges = [...container.querySelectorAll(`.vt-group-badge[data-group-id="${CSS.escape(groupId)}"]`)];
        const allPending = allBadges.every((b) =>
          state.pendingChanges.some((c) => c.type === 'remove' && c.deviceId === b.dataset.deviceId && c.groupId === groupId)
        );
        if (allPending) {
          // All already marked — unmark all
          state.pendingChanges = state.pendingChanges.filter(
            (c) => !(c.type === 'remove' && c.groupId === groupId)
          );
        } else {
          // Mark any not yet staged
          allBadges.forEach((b) => {
            const already = state.pendingChanges.some(
              (c) => c.type === 'remove' && c.deviceId === b.dataset.deviceId && c.groupId === groupId
            );
            if (!already) {
              state.pendingChanges.push({
                type: 'remove',
                deviceId: b.dataset.deviceId,
                deviceName: b.dataset.deviceName,
                groupId,
                groupName,
              });
            }
          });
        }
        renderPendingSummary();
        refreshVehicleTable();
      });
    });
  }
}

function downloadCsv() {
  const statusMap = new Map(state.deviceStatuses.map((s) => [s.device && s.device.id, s]));
  const groupMap  = new Map(state.groups.map((g) => [g.id, g.name || g.id]));

  // Get the same filtered+sorted rows as the current table
  let rows = state.devices.map((d) => getDeviceRow(d, statusMap));
  rows = applySearchFilter(rows, parseSearchQuery(state.vehicleFilterText));
  if (state.columnFilters.plan) rows = rows.filter((r) => r.plan === state.columnFilters.plan);
  if (state.selectedGroupId) {
    const gid = state.selectedGroupId;
    rows = rows.filter((r) => (r.device.groups || []).some((g) => g.id === gid));
  }

  const headers = ['Name', 'Groups', 'Plate', 'Odometer (km)', 'Last Connect', 'First Connect', 'Plan', 'VIN', 'Serial No.', 'Comment'];
  const csvRows = [headers.join(',')];
  for (const r of rows) {
    const d = r.device;
    const groups = (d.groups || []).map((g) => groupMap.get(g.id)).filter(Boolean).join('; ');
    const esc = (v) => `"${String(v == null ? '' : v).replace(/"/g, '""')}"`;
    csvRows.push([d.name, groups, d.licensePlate, r.odoDisplay === '—' ? '' : r.odoDisplay,
      r.lastConnectDisplay, r.firstConnectDisplay, r.plan, d.vehicleIdentificationNumber,
      d.serialNumber, d.comment].map(esc).join(','));
  }

  const blob = new Blob([csvRows.join('\n')], { type: 'text/csv' });
  const url  = URL.createObjectURL(blob);
  const a    = Object.assign(document.createElement('a'), { href: url, download: 'fleet-vehicles.csv' });
  document.body.appendChild(a);
  a.click();
  a.remove();
  URL.revokeObjectURL(url);
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


// =============================================================
// 10. i18n — Apply language and update all text nodes
// =============================================================

// Static map of elementId → string key for text that can be set via textContent
const TEXT_MAP = {
  title:                    'title',
  langToggleBtn:            'langToggle',
  kpiVehiclesLabel:         'kpiVehicles',
  kpiGroupsLabel:           'kpiGroups',
  kpiLicensesLabel:         'kpiLicenses',
  treePanelTitle:           'treePanelTitle',
  loadingText:              'loadingText',
  hideEmptyGroupsLabel:     'hideEmptyGroupsLabel',
  showSystemGroupsLabel:    'showSystemGroupsLabel',
  showArchivedDevicesLabel: 'showArchivedDevicesLabel',
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

  // Reset vehicle panel title when no group selected
  if (!state.selectedGroupId) {
    vehiclePanelTitle.textContent = t('allVehiclesTitle');
  }

  // Re-render charts and lists so labels update
  if (state.groupTree) { renderTreeChart(); renderTreeView(); }
  if (state.devicePlans) renderPlanPills();
  refreshVehicleTable();
}

// =============================================================
// 11. Event Listeners
// =============================================================

langToggleBtn.addEventListener('click', () => {
  const next = state.lang === 'en' ? 'es' : 'en';
  applyLang(next);
});

vehicleSearchInput.addEventListener('input', () => {
  state.vehicleFilterText = vehicleSearchInput.value;
  refreshVehicleTable();
});

document.getElementById('btnDownloadCsv').addEventListener('click', downloadCsv);
document.getElementById('btnColumnConfig').addEventListener('click', function () { openColumnConfigPanel(this); });

document.getElementById('showRelativeDatesChk').addEventListener('change', (e) => {
  state.showRelativeDates = e.target.checked;
  refreshVehicleTable();
});

// ---- Edit mode ----
document.getElementById('btnEditMode').addEventListener('click', () => {
  if (!state.editMode) {
    setEditMode(true);
  } else if (state.pendingChanges.length === 0) {
    setEditMode(false);
  } else {
    openConfirmModal();
  }
});

// Edit group search autocomplete
document.getElementById('editGroupSearch').addEventListener('input', () => {
  const q   = document.getElementById('editGroupSearch').value.toLowerCase();
  const list = document.getElementById('editGroupList');
  if (!q) { list.innerHTML = ''; list.hidden = true; return; }
  const matches = state.groups
    .filter((g) => (g.name || '').toLowerCase().includes(q))
    .slice(0, 20);
  if (!matches.length) { list.innerHTML = ''; list.hidden = true; return; }
  list.innerHTML = matches.map((g) =>
    `<li data-id="${escapeHtml(g.id)}" data-name="${escapeHtml(g.name || g.id)}">${escapeHtml(g.name || g.id)}</li>`
  ).join('');
  list.hidden = false;
});

document.getElementById('editGroupList').addEventListener('click', (e) => {
  const li = e.target.closest('li[data-id]');
  if (!li) return;
  state.editGroupId   = li.dataset.id;
  state.editGroupName = li.dataset.name;
  document.getElementById('editGroupSearch').value = li.dataset.name;
  document.getElementById('editGroupList').hidden = true;
});

document.addEventListener('click', (e) => {
  const list   = document.getElementById('editGroupList');
  const search = document.getElementById('editGroupSearch');
  if (list && search && !search.contains(e.target) && !list.contains(e.target)) {
    list.hidden = true;
  }
});


// Stage Add
document.getElementById('editStageAddBtn').addEventListener('click', () => {
  const status = document.getElementById('editAddStatus');
  const ids    = [...state.editSelectedDevices];
  if (!ids.length)           { setOpStatus(status, 'No devices checked.', true); return; }
  if (!state.editGroupId)    { setOpStatus(status, 'No group selected.', true); return; }

  const existing = state.pendingChanges.find(
    (c) => c.type === 'add' && c.groupId === state.editGroupId
  );
  if (existing) {
    ids.forEach((id) => existing.deviceIds.add(id));
  } else {
    state.pendingChanges.push({
      type: 'add', groupId: state.editGroupId, groupName: state.editGroupName, deviceIds: new Set(ids),
    });
  }
  renderPendingSummary();
  setOpStatus(status, `Staged: ${ids.length} device(s) → "${state.editGroupName}"`);
});

// Stage Create group
document.getElementById('editStageCreateBtn').addEventListener('click', () => {
  const nameEl = document.getElementById('editNewGroupName');
  const status = document.getElementById('editCreateStatus');
  const name   = nameEl.value.trim();
  if (!name) { setOpStatus(status, 'Enter a group name.', true); return; }

  // Temporary ID — replaced with the real Geotab ID on apply
  const tempId = `_pending_${Date.now()}`;
  state.pendingChanges.push({
    type: 'createGroup', name, tempId,
    parentId:   state.editGroupId   || 'GroupCompanyId',
    parentName: state.editGroupName || 'Root',
  });

  // Add immediately to state.groups so it appears in the autocomplete
  // and can be used as target for Add operations in the same edit session
  state.groups.push({ id: tempId, name, children: [], parent: { id: state.editGroupId || 'GroupCompanyId' }, _pending: true });

  nameEl.value = '';
  renderPendingSummary();
  setOpStatus(status, `"${name}" staged — now available in group search`);
});

// Clear all staged + exit edit mode
document.getElementById('editClearAllBtn').addEventListener('click', () => {
  // Remove any temp groups that were added to state for immediate staging
  state.groups = state.groups.filter((g) => !g._pending);
  state.groupTree = buildGroupTree(state.groups, state.devices, state.users, state.showSystemGroups);
  state.pendingChanges = [];
  setEditMode(false);
});

// Confirmation modal
document.getElementById('confirmApply').addEventListener('click', applyPendingChanges);
document.getElementById('confirmCancel').addEventListener('click', () => {
  document.getElementById('confirmModal').hidden = true;
});
document.getElementById('confirmClose').addEventListener('click', () => {
  document.getElementById('confirmModal').hidden = true;
});

// View toggle
const VIEW_BTNS = ['TreeView', 'MindmapView'];
function switchView(activeView) {
  VIEW_BTNS.forEach((v) => {
    const container = document.getElementById(v.charAt(0).toLowerCase() + v.slice(1) + 'Container');
    const btn = document.getElementById('btn' + v);
    if (!container || !btn) return;
    if (v === activeView) {
      container.removeAttribute('hidden');
      btn.classList.add('fo-view-btn--active');
    } else {
      container.setAttribute('hidden', '');
      btn.classList.remove('fo-view-btn--active');
    }
  });
}

document.getElementById('hideEmptyGroupsChk').addEventListener('change', (e) => {
  state.hideEmptyGroups = e.target.checked;
  if (state.groupTree) { renderTreeChart(); renderTreeView(); }
});

document.getElementById('showSystemGroupsChk').addEventListener('change', (e) => {
  state.showSystemGroups = e.target.checked;
  if (state.groups) {
    state.groupTree = buildGroupTree(state.groups, state.devices, state.users, state.showSystemGroups);
    renderTreeChart(); renderTreeView();
  }
});

document.getElementById('showArchivedDevicesChk').addEventListener('change', async (e) => {
  state.showArchivedDevices = e.target.checked;
  if (state.showArchivedDevices && !state.allDevices) {
    // Lazy-load all devices (including archived) on first toggle
    try {
      const all = await apiCall('Get', { typeName: 'Device', resultsLimit: 50000 });
      state.allDevices = all || [];
    } catch (err) {
      state.allDevices = [...state.activeDevices];
    }
  }
  state.devices = state.showArchivedDevices ? (state.allDevices || state.activeDevices) : state.activeDevices;
  state.groupTree = buildGroupTree(state.groups, state.devices, state.users, state.showSystemGroups);
  renderTreeChart(); renderTreeView(); renderKpis(); refreshVehicleTable();
});

document.getElementById('btnTreeView').addEventListener('click', () => switchView('TreeView'));

document.getElementById('btnMindmapView').addEventListener('click', () => {
  switchView('MindmapView');
  renderTreeView();
});


document.getElementById('btnHomeTree').addEventListener('click', () => {
  state.selectedGroupId = null;
  vehiclePanelTitle.textContent = t('allVehiclesTitle');
  // Reset treemap to root view
  if (treeChartInstance) {
    try {
      const root = treeChartInstance.getModel().getSeriesByIndex(0).getData().tree.root;
      treeChartInstance.dispatchAction({ type: 'treemapRootToNode', seriesIndex: 0, targetNode: root });
    } catch (e) { /* ignore */ }
  }
  // Reset mindmap selection
  if (mindmapChartInstance) mindmapChartInstance.dispatchAction({ type: 'unfocusNodeAdjacency' });
  refreshVehicleTable();
});

document.getElementById('btnExpandTree').addEventListener('click', toggleTreePanelExpand);

document.addEventListener('keydown', (e) => {
  if (e.key === 'Escape' && state.treePanelExpanded) toggleTreePanelExpand();
});

// Treemap group search autocomplete
(function () {
  const searchInput = document.getElementById('treeGroupSearch');
  const searchList  = document.getElementById('treeGroupList');
  if (!searchInput || !searchList) return;

  function showTreeSearchResults(text) {
    const q = text.trim().toLowerCase();
    searchList.innerHTML = '';
    if (!q || !state.groups.length) { searchList.hidden = true; return; }
    const matches = state.groups
      .filter((g) => !g._pending && (g.name || g.id).toLowerCase().includes(q))
      .slice(0, 12);
    if (!matches.length) { searchList.hidden = true; return; }
    matches.forEach((g) => {
      const li = document.createElement('li');
      li.textContent = g.name || g.id;
      li.dataset.id  = g.id;
      li.addEventListener('mousedown', (e) => {
        e.preventDefault();
        searchInput.value = '';
        searchList.hidden = true;
        selectGroup(g.id);
        // Navigate treemap step-by-step through the ancestor path to the target node.
        // ECharts treemap requires stepping through each level (direct jumps skip intermediate
        // state and render empty for non-leaf nodes with children).
        if (treeChartInstance && treeChartData) {
          try {
            const internalRoot = treeChartInstance.getModel().getSeriesByIndex(0).getData().tree.root;
            const internalNode = findInternalTreeNode(treeChartData, internalRoot.children, g.id);
            if (internalNode) {
              // Build ancestor path from root down to target (excluding virtual root)
              const path = [];
              let cur = internalNode;
              while (cur && cur.parentNode) { path.unshift(cur); cur = cur.parentNode; }
              // Dispatch each step with a small delay so ECharts processes each transition
              path.forEach((node, i) => {
                setTimeout(() => {
                  treeChartInstance.dispatchAction({ type: 'treemapRootToNode', seriesIndex: 0, targetNode: node });
                }, i * 50);
              });
            }
          } catch (e) { /* internal API unavailable, skip navigation */ }
        }
      });
      searchList.appendChild(li);
    });
    searchList.hidden = false;
  }

  searchInput.addEventListener('input', () => showTreeSearchResults(searchInput.value));
  searchInput.addEventListener('focus', () => showTreeSearchResults(searchInput.value));
  searchInput.addEventListener('blur',  () => { setTimeout(() => { searchList.hidden = true; }, 150); });
  searchInput.addEventListener('keydown', (e) => {
    if (e.key === 'Escape') { searchInput.value = ''; searchList.hidden = true; }
  });
})();

// Admin ops tabs
['Create', 'Assign', 'Remove'].forEach((name) => {
  const btn = document.getElementById('tabBtn' + name);
  const pane = document.getElementById('pane' + name);
  if (!btn || !pane) return;
  btn.addEventListener('click', () => {
    // Deactivate all tabs/panes
    ['Create', 'Assign', 'Remove'].forEach((n) => {
      const b = document.getElementById('tabBtn' + n);
      const p = document.getElementById('pane' + n);
      if (b) b.classList.remove('fo-tab--active');
      if (p) p.setAttribute('hidden', '');
    });
    btn.classList.add('fo-tab--active');
    pane.removeAttribute('hidden');
    state.activeOpsTab = name.toLowerCase();
    refreshVehicleTable();
  });
});

// =============================================================
// 12. Resize Handler
// =============================================================

window.addEventListener('resize', () => {
  if (treeChartInstance)  treeChartInstance.resize();
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
    [treeChartInstance, mindmapChartInstance].forEach((c) => {
      if (c) c.dispose();
    });
    treeChartInstance    = null;
    mindmapChartInstance = null;
  }

  return { initialize, focus, blur };
};
