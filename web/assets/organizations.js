// Organizations (a company's users) on the Settings page, for everyone. Supabase Edge Function API:
// supabase/functions/organizations.
//   Super admins: every organization in a list; create them and assign people.
//   Organization admins: their own organization (people, roles, name; not the auto-join domain).
//   Members: their own organization, read-only, and they may leave it.
// initOrganizations() returns { load }; load() resolves once the organization is shown.

const ROLE_LABELS = { admin: "Admin", member: "Member" };

// Builds an element: el("td", { className: "x" }, "text", child).
function el(tag, props = {}, ...children) {
  const node = Object.assign(document.createElement(tag), props);
  node.append(...children.filter((c) => c !== null && c !== undefined && c !== false));
  return node;
}

const plural = (n, word, many = `${word}s`) => `${n} ${n === 1 ? word : many}`;

function initOrganizations() {
  const $ = (id) => document.getElementById(id);
  const createDialog = $("org-create-dialog");
  const editDialog = $("org-edit-dialog");
  const addDialog = $("org-add-dialog");
  const addForm = $("org-add-form");
  const pickList = $("org-pick-list");

  let superAdmin = false;
  // Super admins: the organization open on the right. Everyone else always sees their own.
  let selectedId = null;
  let current = null; // last GET response
  let allOrgs = [];
  // Whether the caller may change the organization on screen (super admin, or its admin).
  let manage = false;

  async function load() {
    const query = selectedId ? `?org_id=${encodeURIComponent(selectedId)}` : "";
    const { data, error } = await callApi(`organizations${query}`, { method: "GET" });
    if (error) return showMessage(`Could not load organizations: ${error}`, "error");
    superAdmin = !!data.super_admin;
    // Super admins land on the first organization rather than an empty page.
    if (superAdmin && !selectedId && data.organizations?.length) {
      selectedId = data.organizations[0].id;
      return load();
    }
    current = data;
    render();
  }

  function render() {
    const { organizations = [], role } = current;
    manage = !!current.manage;
    $("settings-intro").textContent = superAdmin
      ? "Create organizations and choose who belongs to each. People in the same organization share mailbox settings."
      : role === "admin" ? "Manage who belongs to your organization."
      : "Your organization and the people in it.";
    $("org-layout").hidden = false;
    $("org-layout").classList.toggle("single", !superAdmin);
    $("org-nav-panel").hidden = !superAdmin;
    allOrgs = organizations;
    if (superAdmin) renderNav();

    const org = superAdmin && !selectedId ? null : current.organization;
    $("org-placeholder").hidden = !!org;
    $("org-content").hidden = !org;
    if (!org) {
      $("org-placeholder-text").textContent = superAdmin
        ? "There are no organizations yet. Create one, then add people to it."
        : "You're not in an organization yet. An administrator adds you, or you join automatically when you sign in " +
          "with Microsoft using your company email.";
      $("org-placeholder-new").hidden = !superAdmin;
      return;
    }
    renderDetail(org);
  }

  function renderNav() {
    const text = $("org-search").value.trim().toLowerCase();
    const shown = allOrgs.filter((o) => !text || `${o.name} ${o.domain ?? ""}`.toLowerCase().includes(text));
    $("org-nav").replaceChildren(...shown.map((o) => {
      const b = el("button", { type: "button" },
        el("span", { className: "org-nav-name" }, o.name),
        el("span", { className: "org-nav-meta" }, `${plural(o.member_count, "member")}${o.domain ? ` · @${o.domain}` : ""}`));
      if (o.id === selectedId) b.setAttribute("aria-current", "true");
      b.addEventListener("click", () => {
        selectedId = o.id;
        load();
      });
      return el("li", {}, b);
    }));
    const empty = $("org-nav-empty");
    empty.hidden = shown.length > 0;
    empty.textContent = allOrgs.length ? "No organization matches." : "None yet.";
  }

  function renderDetail(org) {
    const { members = [], invites = [] } = current;
    $("org-name").textContent = org.name;
    $("org-autojoin").replaceChildren(...(org.domain
      ? [el("span", { className: "badge type-ar" }, "Auto-join on"), ` Anyone who signs in with Microsoft using an @${org.domain} address joins as a member.`]
      : [el("span", { className: "badge" }, "Auto-join off"), " People join only when added here."]));
    $("org-count").textContent = `${plural(members.length, "member")}${invites.length ? ` · ${invites.length} invited` : ""}`;
    // Members see the organization read-only; anyone in it may leave.
    $("org-edit").hidden = $("org-add").hidden = $("org-manage-hint").hidden = !manage;
    document.querySelector(".org-danger").hidden = !manage;
    $("org-leave").hidden = !current.role;
    $("org-role-note").textContent = current.role ? `You're ${current.role === "admin" ? "an admin" : "a member"} of this organization.` : "";

    $("org-members").replaceChildren(...members.map(memberRow), ...invites.map(inviteRow));
    $("org-members-empty").hidden = members.length + invites.length > 0;
    $("org-members").closest("table").hidden = members.length + invites.length === 0;
  }

  function memberRow(m) {
    if (!manage) {
      return el("tr", {},
        el("td", {}, m.full_name || "—", m.you ? el("span", { className: "badge you" }, "You") : null),
        el("td", { className: "org-email" }, m.email ?? ""),
        el("td", {}, ROLE_LABELS[m.role]),
        el("td", { className: "org-row-actions" }));
    }
    const select = el("select");
    select.setAttribute("aria-label", `Role of ${m.email}`);
    for (const [value, label] of Object.entries(ROLE_LABELS)) select.add(new Option(label, value, false, value === m.role));
    select.addEventListener("change", () => act({ action: "set_role", user_id: m.user_id, role: select.value },
      `${m.full_name || m.email} is now ${select.value === "admin" ? "an admin" : "a member"}.`));

    let action = null;
    if (!m.you) {
      action = el("button", { type: "button", className: "secondary small" }, "Remove");
      action.addEventListener("click", () => {
        if (confirm(`Remove ${m.full_name || m.email} from ${current.organization.name}?`)) {
          act({ action: "remove_member", user_id: m.user_id }, `Removed ${m.full_name || m.email}.`);
        }
      });
    }
    return el("tr", {},
      el("td", {}, m.full_name || "—", m.you ? el("span", { className: "badge you" }, "You") : null),
      el("td", { className: "org-email" }, m.email ?? ""),
      el("td", {}, select),
      el("td", { className: "org-row-actions" }, action));
  }

  function inviteRow(i) {
    const cancel = el("button", { type: "button", className: "secondary small" }, "Cancel invite");
    cancel.addEventListener("click", () => act({ action: "cancel_invite", email: i.email }, `Cancelled the invitation for ${i.email}.`));
    return el("tr", { className: "org-invite" },
      el("td", {}, el("span", { className: "badge" }, "Invited"), el("div", { className: "hint" }, "Joins on first sign-in")),
      el("td", { className: "org-email" }, i.email),
      el("td", {}, ROLE_LABELS[i.role]),
      el("td", { className: "org-row-actions" }, cancel));
  }

  // Runs an action on the open organization, shows the outcome and reloads.
  async function act(body, done) {
    const { error } = await callApi("organizations", { body: { org_id: current.organization.id, ...body } });
    showMessage(error ?? done, error ? "error" : "ok");
    await load();
    return !error;
  }

  // --- dialogs ---

  for (const b of document.querySelectorAll(".dialog [data-close]")) b.addEventListener("click", () => b.closest("dialog").close());

  function openCreate() {
    $("org-create-form").reset();
    createDialog.showModal();
  }
  $("org-new").addEventListener("click", openCreate);
  $("org-placeholder-new").addEventListener("click", openCreate);
  $("org-search").addEventListener("input", renderNav);

  $("org-create-form").addEventListener("submit", async (e) => {
    e.preventDefault();
    const form = e.currentTarget;
    const { name, domain } = Object.fromEntries(new FormData(form));
    setBusy(form, true);
    const { data, error } = await callApi("organizations", { body: { action: "create", name, domain } });
    setBusy(form, false);
    if (error) return showMessage(error, "error");
    createDialog.close();
    showMessage(`Created ${name}. Now add the people who belong to it.`, "ok");
    selectedId = data.organization.id;
    await load();
    openAdd();
  });

  $("org-edit").addEventListener("click", () => {
    const form = $("org-edit-form");
    form.elements.name.value = current.organization.name;
    form.elements.domain.value = current.organization.domain ?? "";
    // Only super admins change the domain: it decides who joins automatically.
    $("org-edit-domain").hidden = $("org-edit-domain-hint").hidden = !superAdmin;
    editDialog.showModal();
  });

  $("org-edit-form").addEventListener("submit", async (e) => {
    e.preventDefault();
    const form = e.currentTarget;
    const { name, domain } = Object.fromEntries(new FormData(form));
    setBusy(form, true);
    const ok = await act({ action: "update", name, ...(superAdmin ? { domain } : {}) }, "Saved.");
    setBusy(form, false);
    if (ok) editDialog.close();
  });

  $("org-leave").addEventListener("click", async () => {
    const org = current.organization;
    if (!confirm(`Leave ${org.name}? You'll need an admin to add you back.`)) return;
    const me = current.members.find((m) => m.you);
    const { error } = await callApi("organizations", { body: { action: "remove_member", org_id: org.id, user_id: me?.user_id } });
    showMessage(error ?? `You left ${org.name}.`, error ? "error" : "ok");
    if (!error && superAdmin) selectedId = null;
    await load();
  });

  $("org-delete").addEventListener("click", async () => {
    const org = current.organization;
    if (!confirm(`Delete ${org.name}? Everyone in it is taken out of the organization.`)) return;
    const { error } = await callApi(`organizations?org_id=${encodeURIComponent(org.id)}`, { method: "DELETE" });
    if (error) return showMessage(error, "error");
    showMessage(`Deleted ${org.name}.`, "ok");
    document.querySelector(".org-danger").open = false;
    selectedId = null;
    await load();
  });

  // Add people: super admins choose from everyone with an account (and may type addresses of people without one);
  // organization admins type addresses.
  $("org-add").addEventListener("click", openAdd);

  function openAdd() {
    addForm.reset();
    $("org-add-results").hidden = true;
    $("org-add-title").textContent = `Add people to ${current.organization.name}`;
    $("org-pick").hidden = !superAdmin;
    $("org-emails-step").textContent = superAdmin ? "2. Not signed up yet? Enter their email addresses" : "Email addresses";
    $("org-role-step").textContent = superAdmin ? "3. Role" : "Role";
    addDialog.showModal();
    if (superAdmin) {
      $("org-pick-search").focus();
      loadPeople();
    } else {
      addForm.elements.emails.focus();
    }
  }

  async function loadPeople() {
    const status = $("org-pick-status");
    status.textContent = "Loading people…";
    pickList.replaceChildren();
    const orgId = current.organization.id;
    const { data, error } = await callApi("organizations", { body: { action: "list_users", org_id: orgId } });
    if (orgId !== current.organization.id) return;
    if (error) return (status.textContent = error);
    pickList.replaceChildren(...data.users.map((u) => {
      const input = el("input", { type: "checkbox", value: u.email });
      if (u.organization) input.dataset.organization = u.organization;
      const label = el("label", { className: "org-person" }, input,
        el("span", {},
          el("span", { className: "org-person-name" }, u.full_name || u.email),
          el("span", { className: "org-person-meta" }, [u.full_name ? u.email : null, u.organization ? `now in ${u.organization}` : null].filter(Boolean).join(" · "))));
      label.dataset.search = `${u.full_name ?? ""} ${u.email}`.toLowerCase();
      return label;
    }));
    filterPeople();
  }

  function filterPeople() {
    const text = $("org-pick-search").value.trim().toLowerCase();
    const labels = [...pickList.children];
    let shown = 0;
    for (const label of labels) {
      label.hidden = !!text && !label.dataset.search.includes(text);
      if (!label.hidden) shown++;
    }
    const ticked = pickList.querySelectorAll("input:checked").length;
    $("org-pick-status").textContent = !labels.length
      ? "Everyone with an account is already in this organization."
      : `${shown === labels.length ? plural(labels.length, "person", "people") : `${shown} of ${labels.length}`} · ${ticked} selected`;
  }
  $("org-pick-search").addEventListener("input", filterPeople);
  pickList.addEventListener("change", filterPeople);

  addForm.addEventListener("submit", async (e) => {
    e.preventDefault();
    const { emails: typed, role } = Object.fromEntries(new FormData(addForm));
    const ticked = [...pickList.querySelectorAll("input:checked")];
    if (!ticked.length && !typed.trim()) return showMessage(superAdmin ? "Choose people or enter an email address." : "Enter an email address.", "error");
    const moving = ticked.filter((i) => i.dataset.organization);
    if (moving.length && !confirm(`These people will be moved from their current organization:\n\n` +
      `${moving.map((i) => `${i.value} (now in ${i.dataset.organization})`).join("\n")}\n\nContinue?`)) return;

    setBusy(addForm, true);
    const { data, error } = await callApi("organizations", {
      body: { action: "add_members", org_id: current.organization.id, emails: [...ticked.map((i) => i.value), typed].join("\n"), role },
    });
    setBusy(addForm, false);
    if (error) return showMessage(error, "error");

    const failed = data.results.filter((r) => !r.ok);
    const done = data.results.length - failed.length;
    const STATUS = { added: "added", moved: "moved here", invited: "invited" };
    showMessage(`${plural(done, "person", "people")} added to ${current.organization.name}.` +
      (failed.length ? ` ${failed.length} could not be added.` : ""), failed.length ? "error" : "ok");
    await load();
    if (!failed.length) return addDialog.close();
    // Keep the dialog open with what went wrong; failed typed addresses stay to fix and retry.
    $("org-add-results").replaceChildren(...data.results.map((r) =>
      el("li", { className: r.ok ? "ok" : "failed" }, `${r.email}: ${r.ok ? STATUS[r.status] : r.error}`)));
    $("org-add-results").hidden = false;
    const picked = new Set(ticked.map((i) => i.value));
    addForm.elements.emails.value = failed.filter((r) => !picked.has(r.email)).map((r) => r.email).join("\n");
    if (superAdmin) loadPeople();
  });

  return { load };
}
