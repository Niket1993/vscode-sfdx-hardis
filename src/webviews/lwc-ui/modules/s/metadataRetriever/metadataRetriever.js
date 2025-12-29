import { LightningElement, api, track } from "lwc";
import "s/forceLightTheme"; // Ensure light theme is applied


/**
 * Lightning Web Component that retrieves and searches org metadata via a VS Code webview. 
 *
 * Supports two main query modes:
 * - Recent Changes (using SourceMember data)
 * - All Metadata (using Metadata API) 
 *
 * The component communicates with a VS Code extension through window.postMessage-style APIs
 * wrapped by the custom window.sendMessageToVSCode helper. 
 */

// Configuration - Base URL for metadata type documentation
// Modify this URL to change where metadata type links point to
/** @constant {string} Base URL for metadata type documentation links. */
const METADATA_DOC_BASE_URL =
  "https://sf-explorer.github.io/sf-doc-to-json/#/cloud/all/object/";

export default class MetadataRetriever extends LightningElement {
  /**
   * Available Salesforce org connections exposed by the VS Code extension. 
   * Each org entry typically contains alias, username, and instanceUrl fields.
   * @type {Array<Object>}
   * @public
   */
  @api orgs = [];

  /**
   * Available metadata types for the currently selected org.
   * Used to build the metadata type picklist. 
   * @type {Array<{label: string, value: string}>}
   * @public
   */
  @api metadataTypes = [];

/**
   * Username of the currently selected org used for all backend queries. 
   * @type {string|null}
   * @private
   */
  @track selectedOrg = null;

  /**
   * Current query mode: 'recentChanges' (SourceMember-style) or 'allMetadata'. 
   * @type {'recentChanges'|'allMetadata'}
   * @private
   */
  @track queryMode = "recentChanges"; // "recentChanges" or "allMetadata"

  /**
   * Selected metadata type filter; 'All' means no type restriction in recentChanges mode. 
   * @type {string}
   * @private
   */
  @track metadataType = "All";
 /**
   * Installed package filter: 'All', 'Local', or specific namespace. 
   * @type {string}
   * @private
   */
  @track packageFilter = "All";

  /**
   * Options for the package filter picklist, including 'All', 'Local', and any installed namespaces. 
   * @type {Array<{label: string, value: string}>}
   * @private
   */
  @track packageOptions = [
    { label: "All", value: "All" },
    { label: "Local", value: "Local" },
  ];

  /**
   * Client-side metadata name filter used for incremental filtering of loaded records. 
   * @type {string}
   * @private
   */
  @track metadataName = "";

  /**
   * Client-side filter for the last modified by user name (substring match). 
   * @type {string}
   * @private
   */
  @track lastUpdatedBy = "";

  /**
   * Start date filter (inclusive) for last modified date in recentChanges mode. 
   * @type {string}
   * @private
   */
  @track dateFrom = "";

  /**
   * End date filter (inclusive) for last modified date in recentChanges mode. 
   * @type {string}
   * @private
   */
  @track dateTo = "";
 /**
   * Free-text search term applied across type, name, and user fields in the current result set. 
   * @type {string}
   * @private
   */
  @track searchTerm = "";

  /**
   * Indicates whether at least one server-side search has been performed. 
   * @type {boolean}
   * @private
   */
  @track hasSearched = false;


 /**
   * Controls whether the backend annotates results with local file existence information. 
   * @type {boolean}
   * @private
   */
  @track checkLocalFiles = false;

  /**
   * Indicates that local file checks are supported in the current workspace
   * (sfdx-project.json found at the workspace root). 
   * @type {boolean}
   * @private
   */
  @track checkLocalAvailable = true;

  /**
   * True while org list is being loaded from the extension backend. 
   * @type {boolean}
   * @private
   */
  @track isLoadingOrgs = false;

  /**
   * True while installed package information is being loaded for the selected org. 
   * @type {boolean}
   * @private
   */
  @track isLoadingPackages = false;

 /**
   * True while a metadata query is in progress. 
   * @type {boolean}
   * @private
   */
  @track isLoading = false;

  /**
   * Raw metadata records returned from the backend (before client-side filters). 
   * @type {Array<Object>}
   * @private
   */
  @track metadata = [];

  /**
   * Metadata records after applying all client-side filters and search terms. 
   * @type {Array<Object>}
   * @private
   */
  @track filteredMetadata = [];

  /**
   * Last error message from a failed query or backend call, if any. 
   * @type {string|null}
   * @private
   */
  @track error = null;

  /**
   * Full row objects for all currently selected metadata items across filters. 
   * @type {Array<Object>}
   * @private
   */
  @track selectedRows = [];


/**
   * Unique keys of selected rows (MemberType::MemberName), kept stable across filtering and sorting. 
   * @type {Array<string>}
   * @private
   */
  @track selectedRowKeys = [];

  /**
   * Indicates whether the hidden feature modal (Masha Easter egg) is visible. 
   * @type {boolean}
   * @private
   */
  @track showFeature = false;

  /**
   * Randomized HTML id used to uniquely identify the Easter egg modal instance. 
   * @type {string|null}
   * @private
   */
  @track featureId = null;

  /**
   * Text displayed in the Easter egg modal dialog. 
   * @type {string}
   * @private
   */
  @track featureText;

  /**
   * Image URL for the feature logo displayed inside the Easter egg modal. 
   * @type {string}
   * @private
   */
  @track imgFeatureLogo = "";

 // Local package selector (sfdx-project.json packageDirectories)

  /**
   * Local package directory options derived from sfdx-project.json. 
   * @type {Array<{label: string, value: string}>}
   * @private
   */
  @track localPackageOptions = [];

  /**
   * Currently selected local package directory used for retrieve operations. 
   * @type {string|null}
   * @private
   */
  @track selectedLocalPackage = null;

  /**
   * Initial local package value to restore if the selector becomes disabled or reset. 
   * @type {string|null}
   * @private
   */
  @track initialLocalPackage = null;

  /**
   * Indicates whether one or more retrieve operations are currently running. 
   * @type {boolean}
   * @private
   */
  @track isRetrieving = false;

  // Performance optimization properties

  /**
   * Timer id used to debounce search-in-results input to reduce filter recalculations. 
   * @type {number|null}
   * @private
   */
  searchDebounceTimer = null;

  /**
   * Cached normalized "from" date object for efficient date filtering. 
   * @type {Date|null}
   * @private
   */
  cachedDateFrom = null;

  /**
   * Cached normalized "to" date object for efficient date filtering. 
   * @type {Date|null}
   * @private
   */
  cachedDateTo = null;






/**
   * Inverse of hasSearched, useful for template conditions that show pre-search state. 
   * @type {boolean}
   * @readonly
   */
  get hasSearchedd() {
    return !this.hasSearched;
  }


/**
   * Dynamically computed Lightning datatable column configuration
   * that adapts to the active query mode and local file toggle. 
   *
   * In recentChanges mode it adds an Operation emoji column,
   * and when local checks are enabled it appends a Local column. 
   *
   * @type {Array<Object>}
   * @readonly
   */
  get columns() {
    // Build columns step by step so we can insert the Change icon column
    const cols = [];

    // If in Recent Changes mode, insert change icon column after Metadata Name
    if (this.isRecentChangesMode) {
      // Emoji column for change operation (created/modified/deleted)
      cols.push({
        label: "Operation",
        fieldName: "ChangeIcon",
        type: "text",
        cellAttributes: {
          alignment: "center",
        },
        initialWidth: 30,
      });
    }

    // Metadata Type
    cols.push({
      label: "Metadata Type",
      fieldName: "MemberTypeUrl",
      type: "url",
      sortable: true,
      wrapText: true,
      initialWidth: 160,
      typeAttributes: {
        label: { fieldName: "MemberType" },
        tooltip: { fieldName: "MemberTypeTitle" },
        target: "_blank",
      },
    });

    // Metadata Name
    cols.push({
      label: "Metadata Name",
      fieldName: "MemberName",
      type: "button",
      sortable: true,
      wrapText: true,
      typeAttributes: {
        label: { fieldName: "MemberName" },
        title: { fieldName: "MemberNameTitle" },
        name: "open",
        variant: "base",
      },
      cellAttributes: {
        alignment: "left",
        class: "metadata-name-button",
      },
    });

    // Last Updated By
    cols.push({
      label: "Last Updated By",
      fieldName: "LastModifiedByName",
      type: "text",
      sortable: true,
      wrapText: true,
      initialWidth: 165,
    });

    // Last Updated Date
    cols.push({
      label: "Last Updated Date",
      fieldName: "LastModifiedDate",
      type: "date",
      sortable: true,
      typeAttributes: {
        year: "numeric",
        month: "short",
        day: "2-digit",
        hour: "2-digit",
        minute: "2-digit",
      },
      initialWidth: 165,
    });

    // Local file existence column (centered) - only when the user enabled the toggle
    if (this.checkLocalFiles) {
      cols.push({
        label: "Local",
        fieldName: "LocalFileIcon",
        type: "text",
        cellAttributes: {
          alignment: "center",
        },
        initialWidth: 30,
      });
    }

    // Add download icon column (single-icon button)
    // Use a 'button-icon' column so users can click the download icon directly
    cols.push({
      type: "button-icon",
      typeAttributes: {
        iconName: "utility:download",
        title: "Download",
        variant: "bare",
        alternativeText: "Download",
        name: "download",
      },
      initialWidth: 30,
      cellAttributes: {
        alignment: "center",
      },
    });

    return cols;
  }


  
  /**
   * Org picklist options derived from the orgs array, sorted alphabetically by label. 
   *
   * The label prefers instanceUrl (normalized) and falls back to alias or username. 
   *
   * @type {Array<{label: string, value: string}>}
   * @readonly
   */

  get orgOptions() {
    if (!this.orgs || !Array.isArray(this.orgs)) {
      return [];
    }
    const formatLabel = (org) => {
      if (org.instanceUrl) {
        return org.instanceUrl
          .replace(/^https?:\/\//i, "")
          .replace(/\/$/, "")
          .replace(/\.my\.salesforce\.com$/i, "");
      }
      return org.alias || org.username;
    };

    const sortedOrgs = [...this.orgs].sort((a, b) => {
      const nameA = formatLabel(a).toLowerCase();
      const nameB = formatLabel(b).toLowerCase();
      return nameA.localeCompare(nameB);
    });

    const sortedOrgsValues = sortedOrgs.map((org) => ({
      label: formatLabel(org),
      value: org.username,
    }));

    if (this.isLoadingOrgs) {
      sortedOrgsValues.push({ label: "Loading...", value: "" });
    }

    return sortedOrgsValues;
  }


  /**
   * Metadata type picklist options depending on query mode. 
   *
   * In allMetadata mode, the synthetic 'All' option is omitted
   * to force the user to choose a specific type. 
   *
   * @type {Array<{label: string, value: string}>}
   * @readonly
   */

  get metadataTypeOptions() {
    // In All Metadata mode, don't include "All" option
    const options =
      this.queryMode === "allMetadata" ? [] : [{ label: "All", value: "All" }];
    if (this.metadataTypes && Array.isArray(this.metadataTypes)) {
      return options.concat(this.metadataTypes);
    }
    return options;
  }


   /**
   * Query mode picklist options shown in the UI. 
   * @type {Array<{label: string, value: string}>}
   * @readonly
   */

  get queryModeOptions() {
    return [
      { label: "Recent Changes", value: "recentChanges" },
      { label: "All Metadata", value: "allMetadata" },
    ];
  }


  /**
   * Convenience flag indicating that the component is in recentChanges mode. 
   * @type {boolean}
   * @readonly
   */

  get isRecentChangesMode() {
    return this.queryMode === "recentChanges";
  }


  /**
   * Convenience flag indicating that the component is in allMetadata mode. 
   * @type {boolean}
   * @readonly
   */

  get isAllMetadataMode() {
    return this.queryMode === "allMetadata";
  }

   /**
   * True when the "Check local files" toggle should be disabled
   * because the workspace does not support local checks. 
   * @type {boolean}
   * @readonly
   */
  
  get checkLocalDisabled() {
    return !this.checkLocalAvailable;
  }


  /**
   * Tooltip text for the "Check local files" toggle that explains why it is disabled. 
   * @type {string}
   * @readonly
   */

  get checkLocalTooltip() {
    return this.checkLocalAvailable
      ? ""
      : "No sfdx-project.json found in workspace root - local file checks disabled";
  }


   /**
   * Indicates whether there are any filtered results to render in the datatable. 
   * @type {boolean}
   * @readonly
   */

  get hasResults() {
    return this.filteredMetadata && this.filteredMetadata.length > 0;
  }

    /**
   * Whether the results area (table, search in results, and bulk actions)
   * should be shown. 
   *
   * It remains visible as long as any metadata has been loaded from the backend,
   * even if client-side filters are hiding all rows. 
   *
   * @type {boolean}
   * @readonly
   */
  get showResultsArea() {
    return this.hasResults || this.hasMetadataLoaded;
  }


/**
   * True when any metadata has been loaded from the backend regardless of filtering. 
   * @type {boolean}
   * @readonly
   */
  get hasMetadataLoaded() {
    return this.metadata && this.metadata.length > 0;
  }


/**
   * True when a search has been performed and there are no rows to show,
   * and the component is not currently loading. 
   * @type {boolean}
   * @readonly
   */

  get noResults() {
    // Only show the No Results state when a search has been performed
    return (
      this.hasSearched &&
      !this.isLoading &&
      this.filteredMetadata &&
      this.filteredMetadata.length === 0
    );
  }


  /**
   * Indicates that the component currently holds an error message. 
   * @type {boolean}
   * @readonly
   */

  get hasError() {
    return this.error !== null;
  }


  
  /**
   * Indicates whether the current input state is sufficient to perform a search. 
   *
   * In allMetadata mode a specific metadata type is required,
   * while recentChanges mode can run with type = 'All'. 
   *
   * @type {boolean}
   * @readonly
   */

  get canSearch() {
    if (this.selectedOrg === null) {
      return false;
    }
    // In All Metadata mode, require a specific metadata type
    if (
      this.queryMode === "allMetadata" &&
      (this.metadataType === "All" || !this.metadataType)
    ) {
      return false;
    }
    return true;
  }


   /**
   * Negation helper for canSearch, used for disabled states in the template. 
   * @type {boolean}
   * @readonly
   */

  get cannotSearch() {
    return !this.canSearch;
  }

   /**
   * True when at least one row is selected across all metadata. 
   * @type {boolean}
   * @readonly
   */

  get hasSelectedRows() {
    return this.selectedRows && this.selectedRows.length > 0;
  }


  /**
   * Label for the Retrieve Selected button that includes the number of selected items. 
   * @type {string}
   * @readonly
   */

  get retrieveSelectedLabel() {
    const count = this.selectedRows ? this.selectedRows.length : 0;
    return count > 0
      ? `Retrieve ${count} Selected Metadata`
      : "Retrieve Selected Metadata";
  }

  
  /**
   * Lifecycle hook that runs when the component is inserted into the DOM. 
   *
   * Initializes org loading, sets up scroll and resize listeners
   * to manage the floating retrieve button visibility, and triggers
   * an initial visibility check after first render. 
   *
   * @override
   */

  connectedCallback() {
    // Notify VS Code that the component is initialized
    this.isLoadingOrgs = true;
    window.sendMessageToVSCode({ type: "listOrgs" });
    // Bind a debounced visibility check for the floating retrieve button
    this._visibilityDebounceTimer = null;
    this._boundDoDebouncedCheck = () => {
      // Debounce: wait a small time before performing the expensive DOM checks
      if (this._visibilityDebounceTimer) {
        clearTimeout(this._visibilityDebounceTimer);
      }
      this._visibilityDebounceTimer = setTimeout(() => {
        this.checkRetrieveButtonVisibility();
        this._visibilityDebounceTimer = null;
      }, 120); // 120ms debounce to avoid layout thrashing during scroll/resize
    };

    // Listen to global scroll/resize events to detect when the main button goes off-screen
    window.addEventListener("scroll", this._boundDoDebouncedCheck, true);
    window.addEventListener("resize", this._boundDoDebouncedCheck);
    // Initial evaluation after the UI has rendered
    setTimeout(() => this._boundDoDebouncedCheck(), 50);
  }



   /**
 * Lifecycle hook invoked when the component is removed from the DOM. [web:14]
 *
 * Cleans up global scroll and resize listeners used to track the visibility of the
 * main retrieve button, and clears any pending debounce timer to prevent memory
 * leaks or callbacks executing after the component is destroyed. [web:20][web:31]
 *
 * @override
 */

  disconnectedCallback() {
    // Clean up listeners
    if (this._boundDoDebouncedCheck) {
      window.removeEventListener("scroll", this._boundDoDebouncedCheck, true);
      window.removeEventListener("resize", this._boundDoDebouncedCheck);
      this._boundDoDebouncedCheck = null;
    }
    if (this._visibilityDebounceTimer) {
      clearTimeout(this._visibilityDebounceTimer);
      this._visibilityDebounceTimer = null;
    }
  }


/**
 * Initializes the component state from data provided by the VS Code extension host. [web:21][web:30]
 *
 * This method:
 * - Populates the available orgs and selects a default org.
 * - Loads metadata types for the selected org.
 * - Configures whether local file checks are available.
 * - Initializes local package options from sfdx-project.json metadata. [web:16][web:25]
 *
 * @public
 * @param {Object} data Initialization payload from the host extension.
 * @param {Array<Object>} [data.orgs] List of org connection descriptors (alias, username, instanceUrl, etc.).
 * @param {string} [data.selectedOrgUsername] Username of the org that should be selected by default.
 * @param {Array<{label:string,value:string}>} [data.metadataTypes] Initial list of metadata type options.
 * @param {boolean} [data.checkLocalAvailable] Whether local file existence checks are supported for this workspace.
 * @param {Array<{label:string,value:string}>} [data.localPackageOptions] Local package directory options from sfdx-project.json.
 * @param {string} [data.defaultLocalPackage] Default local package to target for retrieve operations.
 */

  @api
  initialize(data) {
    if (data) {
      if (data.orgs && Array.isArray(data.orgs)) {
        this.orgs = data.orgs;
        // Set default org if provided or use first available
        if (data.selectedOrgUsername) {
          this.selectedOrg = data.selectedOrgUsername;
        } else if (this.orgs.length > 0) {
          this.selectedOrg = this.orgs[0].username;
        }
        window.sendMessageToVSCode({
          type: "listMetadataTypes",
          data: { username: this.selectedOrg },
        });
      }
      if (data.metadataTypes && Array.isArray(data.metadataTypes)) {
        this.metadataTypes = data.metadataTypes;
      }
      // Backend can indicate whether local file checking is available
      if (typeof data.checkLocalAvailable === "boolean") {
        this.checkLocalAvailable = data.checkLocalAvailable;
        if (!this.checkLocalAvailable) {
          this.checkLocalFiles = false; // force-uncheck
        }
      }

      // Local packages from sfdx-project.json
      if (data.localPackageOptions && Array.isArray(data.localPackageOptions)) {
        this.localPackageOptions = data.localPackageOptions;
      }
      if (data.defaultLocalPackage) {
        this.selectedLocalPackage = data.defaultLocalPackage;
        this.initialLocalPackage = data.defaultLocalPackage;
      } else if (
        this.localPackageOptions &&
        Array.isArray(this.localPackageOptions) &&
        this.localPackageOptions.length > 0
      ) {
        this.selectedLocalPackage = this.localPackageOptions[0].value;
        this.initialLocalPackage = this.selectedLocalPackage;
      }
    }
  }

  

  /**
 * Indicates whether the local package selector should be shown in the UI. [web:31]
 *
 * The selector is displayed only when:
 * - At least one metadata row is selected.
 * - More than one local package option is available. [web:31]
 *
 * @type {boolean}
 * @readonly
 */

  get showLocalPackageSelector() {
    return (
      this.hasSelectedRows &&
      this.localPackageOptions &&
      Array.isArray(this.localPackageOptions) &&
      this.localPackageOptions.length > 1
    );
  }


  /**
 * True when the local package selector must be disabled, typically while
 * a retrieve operation is in progress to prevent target package changes. [web:31]
 *
 * @type {boolean}
 * @readonly
 */

  get localPackageDisabled() {
    return this.isRetrieving === true;
  }

  
/**
 * Handles changes to the local package selector and updates the currently
 * targeted SFDX package directory for retrieve operations. [web:31]
 *
 * Changes are ignored while a retrieve is in progress. [web:31]
 *
 * @param {CustomEvent} event lightning-combobox change event containing the selected package value.
 */

  handleLocalPackageChange(event) {
    if (this.isRetrieving === true) {
      return;
    }
    this.selectedLocalPackage = event.detail.value;
  }

  
/**
 * Handles org selection changes, clears existing state, and triggers loading
 * of packages and metadata types for the newly selected org. [web:16][web:31]
 *
 * @param {CustomEvent} event lightning-combobox change event with the selected org username.
 */

  handleOrgChange(event) {
    this.selectedOrg = event.detail.value;
    // When org changes, clear any current results and selections
    this.metadata = [];
    this.filteredMetadata = [];
    this.selectedRows = [];
    this.selectedRowKeys = [];
    this.error = null;
    this.isLoading = false;
    // Reset search state when switching orgs
    this.hasSearched = false;

    // When org changes, lazy-load installed package namespaces for that org
    this.isLoadingPackages = true;
    window.sendMessageToVSCode({
      type: "listPackages",
      data: { username: this.selectedOrg },
    });
    // Reset package filter to All
    this.packageFilter = "All";
    // When org changes, lazy-load available metadatas for that org
    window.sendMessageToVSCode({
      type: "listMetadataTypes",
      data: { username: this.selectedOrg },
    });

  }

  /**
 * Handles toggle of the "Check local files" option and optionally re-runs
 * the server query to annotate results with local file existence data. [web:16][web:31]
 *
 * @param {CustomEvent} event lightning-input change event for the toggle.
 */

  handleCheckLocalChange(event) {
    if (!this.checkLocalAvailable) {
      this.checkLocalFiles = false;
      return;
    }
    // lightning-input toggle may expose the boolean on event.target.checked or event.detail.checked
    const newVal =
      event.target?.checked === true ||
      event.detail?.checked === true ||
      event.detail?.value === true;
    this.checkLocalFiles = newVal;

    // If the user just enabled the toggle and we already performed a search,
    // re-run the server query to request annotated results (LocalFileExists).
    if (newVal === true && this.hasSearched && this.canSearch) {
      // small timeout to allow UI to update toggle state before triggering search
      setTimeout(() => this.handleSearch(), 50);
    }
  }


  /**
 * Handles changes to the query mode (recentChanges vs allMetadata),
 * resetting incompatible filters and clearing results accordingly. [web:31]
 *
 * @param {CustomEvent} event lightning-combobox change event for query mode.
 */

  handleQueryModeChange(event) {
    this.queryMode = event.detail.value;
    // Reset metadata type when switching to All Metadata mode (force user to select)
    if (this.queryMode === "allMetadata") {
      if (this.metadataType === "All") {
        this.metadataType = "";
      }
    } else {
      // Reset to "All" when switching back to Recent Changes mode
      if (!this.metadataType) {
        this.metadataType = "All";
      }
    }
    // Clear results when switching modes
    this.metadata = [];
    this.filteredMetadata = [];
    this.selectedRows = [];
    this.selectedRowKeys = [];
    // Reset search state when switching query modes
    this.hasSearched = false;
  }


  
/**
 * Handles row selection changes from the datatable, maintaining a global
 * list of selected rows that persists across filtering and sorting. [web:31]
 *
 * @param {CustomEvent} event lightning-datatable onrowselection event.
 */

  handleRowSelection(event) {
    const currentlySelectedRows = event.detail.selectedRows;
    const currentlySelectedKeys = currentlySelectedRows.map(
      (row) => row.uniqueKey,
    );

    // Get keys of currently visible rows in the datatable
    const visibleKeys = this.filteredMetadata.map((row) => row.uniqueKey);

    // Remove unselected visible keys from master list
    this.selectedRowKeys = this.selectedRowKeys.filter(
      (key) => !visibleKeys.includes(key),
    );

    // Add newly selected keys
    this.selectedRowKeys = [...this.selectedRowKeys, ...currentlySelectedKeys];

    // Update selectedRows to include all selected items from metadata (not just filtered)
    this.selectedRows = this.metadata.filter((row) =>
      this.selectedRowKeys.includes(row.uniqueKey),
    );

    // Update floating retrieve button visibility after selection changes
    // Use a timeout to ensure DOM updates are applied before measuring
    setTimeout(() => this.checkRetrieveButtonVisibility(), 0);
  }


  /**
 * Handles metadata type filter changes and reapplies client-side filters. [web:31]
 *
 * @param {CustomEvent} event lightning-combobox change event.
 */

  handleMetadataTypeChange(event) {
    this.metadataType = event.detail.value;
    this.applyFilters();
  }


  
/**
 * Handles package filter changes (All, Local, or namespace) and reapplies filters. [web:31]
 *
 * @param {CustomEvent} event lightning-combobox change event.
 */

  handlePackageChange(event) {
    this.packageFilter = event.detail.value;
    // Update filters immediately
    this.applyFilters();
  }


  /**
 * Handles metadata name text input changes and reapplies filters using a
 * case-insensitive substring match against MemberName. [web:31]
 *
 * @param {Event} event Input change event from the name textbox.
 */

  handleMetadataNameChange(event) {
    this.metadataName = event.target.value;
    this.applyFilters();
  }


  /**
 * Handles last updated by text input changes, triggers a hidden Easter egg
 * when the user enters "Masha", and reapplies filters. [web:31]
 *
 * @param {Event} event Input change event from the user textbox.
 */

  handleLastUpdatedByChange(event) {
    this.lastUpdatedBy = event.target.value;
    // Easter egg: show modal when user types 'Masha' (case-insensitive)
    try {
      const v = (this.lastUpdatedBy || "").toString().trim();
      if (v.toLowerCase() === "masha") {
        // random feature id for element attributes
        this.featureId = Math.random().toString(36).slice(2, 10);
        // Calculate number of days before November 29, 2025
        const days =
          Math.ceil(
            (new Date("2025-11-29") - new Date()) / (1000 * 60 * 60 * 24),
          ) - 1;
        this.featureText = `See you in ${days} days 😘`;
        this.showFeature = true;
        // Add keydown listener to close on ESC
        this._boundFeatureKeydown = (e) => {
          if (e.key === "Escape") {
            this.hideFeature();
          }
        };
        window.addEventListener("keydown", this._boundFeatureKeydown);
      }
    } catch (e) {
      // ignore
    }

    this.applyFilters();
  }

  
/**
 * Hides the  modal and removes the Escape key listener
 * that allows closing the modal from the keyboard. [web:31]
 */

  hideFeature() {
    this.showFeature = false;
    this.featureId = null;
    if (this._boundFeatureKeydown) {
      window.removeEventListener("keydown", this._boundFeatureKeydown);
      this._boundFeatureKeydown = null;
    }
  }

  /**
 * Handles From date changes for the date filter and reapplies filters. [web:31]
 *
 * @param {Event} event Date input change event.
 */

  handleDateFromChange(event) {
    this.dateFrom = event.target.value;
    this.applyFilters();
  }

  /**
 * Handles To date changes for the date filter and reapplies filters. [web:31]
 *
 * @param {Event} event Date input change event.
 */

  handleDateToChange(event) {
    this.dateTo = event.target.value;
    this.applyFilters();
  }

  /**
 * Handles search-in-results text input with a debounce timer to reduce
 * the number of filter recalculations and maintain selection state. [web:31]
 *
 * @param {Event} event Input change event from the search textbox.
 */

  handleSearchChange(event) {
    this.searchTerm = event.target.value;
    // Debounce the filter application
    if (this.searchDebounceTimer) {
      clearTimeout(this.searchDebounceTimer);
    }
    this.searchDebounceTimer = setTimeout(() => {
      this.applyFilters();
      // Force re-render of datatable to restore selection state
      this.selectedRowKeys = [...this.selectedRowKeys];
      this.searchDebounceTimer = null;
    }, 300);
  }

  /**
 * Executes a server-side metadata query using the current filter criteria
 * and query mode via the VS Code extension messaging API. [web:16][web:21]
 */

  handleSearch() {
    if (!this.canSearch) {
      return;
    }

    this.isLoading = true;
    this.error = null;
    // Mark that a search has been performed
    this.hasSearched = true;
    // Clear existing metadata and client-side text filter to ensure a fresh server search
    this.metadata = [];
    this.filteredMetadata = [];
    // Clear client-side search term and any pending debounce
    this.searchTerm = "";
    if (this.searchDebounceTimer) {
      clearTimeout(this.searchDebounceTimer);
      this.searchDebounceTimer = null;
    }
    this.selectedRows = [];
    this.selectedRowKeys = [];

    // Send filter criteria to VS Code with query mode
    window.sendMessageToVSCode({
      type: "queryMetadata",
      data: {
        username: this.selectedOrg,
        queryMode: this.queryMode,
        metadataType:
          this.metadataType && this.metadataType !== "All"
            ? this.metadataType
            : null,
        metadataName: this.metadataName || null,
        packageFilter:
          this.packageFilter && this.packageFilter !== "All"
            ? this.packageFilter
            : null,
        lastUpdatedBy: this.isRecentChangesMode
          ? this.lastUpdatedBy || null
          : null,
        dateFrom: this.isRecentChangesMode ? this.dateFrom || null : null,
        dateTo: this.isRecentChangesMode ? this.dateTo || null : null,
        checkLocalFiles: this.checkLocalFiles || false,
      },
    });
  }


  /**
 * Sends the currently selected metadata items to the extension for bulk retrieval
 * into the selected local package directory. [web:16][web:21]
 */

  handleRetrieveSelected() {
    if (!this.hasSelectedRows) {
      return;
    }

    // Send selected metadata to VS Code for bulk retrieval
    window.sendMessageToVSCode({
      type: "retrieveSelectedMetadata",
      data: {
        username: this.selectedOrg,
        localPackage: this.selectedLocalPackage,
        metadata: this.selectedRows.map((row) => ({
          memberType: row.MemberType,
          memberName: row.MemberName,
          deleted: row.ChangeIcon === "🔴",
        })),
      },
    });
  }

  /**
 * Resets all client-side filters and selection, while keeping
 * the existing metadata records loaded from the last query. [web:31]
 */

  handleClearFilters() {
    this.metadataType = "All";
    this.metadataName = "";
    this.lastUpdatedBy = "";
    this.dateFrom = "";
    this.dateTo = "";
    this.searchTerm = "";
    this.packageFilter = "All";
    this.selectedRows = [];
    this.selectedRowKeys = [];
    this.applyFilters();
  }


  /**
 * Requests the VS Code extension to open the local folder
 * where retrieved metadata is stored (retrieve history). [web:16]
 */

  handleViewHistory() {
    window.sendMessageToVSCode({
      type: "openRetrieveFolder",
      data: {},
    });
  }

  
/**
 * Applies all client-side filters (type, name, user, date range, package, search)
 * to the metadata array and updates filteredMetadata in a single pass. [web:31]
 */

  applyFilters() {
    if (!this.metadata || this.metadata.length === 0) {
      this.filteredMetadata = [];
      return;
    }

    // Cache date objects to avoid creating new ones for every item
    let fromDate = null;
    if (this.dateFrom) {
      fromDate = new Date(this.dateFrom);
      if (!isNaN(fromDate.getTime())) {
        fromDate.setHours(0, 0, 0, 0);
      } else {
        fromDate = null;
      }
    }

    let toDate = null;
    if (this.dateTo) {
      toDate = new Date(this.dateTo);
      if (!isNaN(toDate.getTime())) {
        toDate.setHours(23, 59, 59, 999);
      } else {
        toDate = null;
      }
    }

    // Cache lowercase strings to avoid multiple toLowerCase() calls
    const metadataNameLower = this.metadataName
      ? this.metadataName.toLowerCase()
      : null;
    const userLower = this.lastUpdatedBy
      ? this.lastUpdatedBy.toLowerCase()
      : null;
    const searchLower = this.searchTerm ? this.searchTerm.toLowerCase() : null;

    // Single pass filtering
    this.filteredMetadata = this.metadata.filter((item) => {
      // Apply metadata type filter
      if (this.metadataType && this.metadataType !== "All") {
        if (item.MemberType !== this.metadataType) {
          return false;
        }
      }

      // Apply metadata name filter
      if (metadataNameLower) {
        if (
          !item.MemberName ||
          !item.MemberName.toLowerCase().includes(metadataNameLower)
        ) {
          return false;
        }
      }

      // Apply last updated by filter
      if (userLower) {
        if (
          !item.LastModifiedByName ||
          !item.LastModifiedByName.toLowerCase().includes(userLower)
        ) {
          return false;
        }
      }

      // Apply date range filters
      if (fromDate && item.LastModifiedDate) {
        const itemDate = new Date(item.LastModifiedDate);
        if (itemDate < fromDate) {
          return false;
        }
      }

      if (toDate && item.LastModifiedDate) {
        const itemDate = new Date(item.LastModifiedDate);
        if (itemDate > toDate) {
          return false;
        }
      }

      // Apply package filter (client-side)
      if (this.packageFilter && this.packageFilter !== "All") {
        const pf = this.packageFilter;
        const fullName = item.MemberName || "";
        const compName = fullName.includes(".")
          ? fullName.split(".").pop() || fullName
          : fullName;
        if (pf === "Local") {
          // Local = ends with official suffix AND has only one __ (no namespace prefix)
          const doubleUnderscoreCount = (compName.match(/__/g) || []).length;

          if (doubleUnderscoreCount === 0) {
            // No __ at all -> local (standard metadata)
            return true;
          }

          if (doubleUnderscoreCount === 1) {
            // One __: check if it's an official suffix
            const officialSuffixes = [
              "__c",
              "__r",
              "__x",
              "__s",
              "__mdt",
              "__b",
            ];
            const hasOfficialSuffix = officialSuffixes.some((suffix) =>
              compName.endsWith(suffix),
            );
            if (hasOfficialSuffix) {
              return true; // local
            }
            // One __ but no official suffix (e.g., CodeBuilder__something) -> packaged
            return false;
          }

          // Multiple __ -> packaged
          return false;
        } else {
          // Component segment must start with namespace__ pattern
          const nsPattern = `${pf}__`;
          if (!compName.startsWith(nsPattern)) {
            return false; // not this namespace -> exclude
          }
        }
      }

      // Apply search term filter (searches across all fields)
      if (searchLower) {
        const matchesSearch =
          (item.MemberType &&
            item.MemberType.toLowerCase().includes(searchLower)) ||
          (item.MemberName &&
            item.MemberName.toLowerCase().includes(searchLower)) ||
          (item.LastModifiedByName &&
            item.LastModifiedByName.toLowerCase().includes(searchLower));
        if (!matchesSearch) {
          return false;
        }
      }

      return true;
    });
  }


  /**
 * Handles row-level actions from the datatable, such as:
 * - 'download': retrieve a single metadata item.
 * - 'open': request opening the local metadata file in VS Code. [web:16]
 *
 * @param {CustomEvent} event lightning-datatable onrowaction event.
 */

  handleRowAction(event) {
    // datatable action events provide event.detail.action (with a name)
    // button-icon columns provide event.detail.name directly
    const row = event.detail.row || event.detail.payload;
    const actionName =
      (event.detail.action && event.detail.action.name) ||
      event.detail.name ||
      null;

    if (actionName === "download") {
      // support legacy 'retrieve' and new 'download' name
      this.handleRetrieve(row);
      return;
    }

    if (actionName === "open") {
      // user clicked the metadata name button -> request extension to open file
      window.sendMessageToVSCode({
        type: "openMetadataFile",
        data: { metadataType: row.MemberType, metadataName: row.MemberName },
      });
      return;
    }
  }

  /**
 * Sends a request to retrieve a single metadata item into the selected local package. [web:16]
 *
 * @param {Object} row Row object with MemberType and MemberName fields.
 */

  handleRetrieve(row) {
    window.sendMessageToVSCode({
      type: "retrieveMetadata",
      data: {
        username: this.selectedOrg,
        localPackage: this.selectedLocalPackage,
        memberType: row.MemberType,
        memberName: row.MemberName,
        deleted: row.ChangeIcon === "🔴",
      },
    });
  }


  /**
 * Main message entry point for events sent from the VS Code extension
 * into this webview. [web:16][web:21]
 *
 * Dispatches messages to specialized handlers such as initialize, queryResults, etc. [web:31]
 *
 * @public
 * @param {string} type Message type identifier.
 * @param {Object} data Associated payload for the message.
 */

  @api
  handleMessage(type, data) {
    if (type === "initialize") {
      this.initialize(data);
    } else if (type === "imageResources") {
      this.handleImageResources(data);
    } else if (type === "listOrgsResults") {
      this.handleOrgResults(data);
    } else if (type === "listPackagesResults") {
      this.handleListPackagesResults(data);
    } else if (type === "listMetadataTypesResults") {
      this.handleListMetadataTypesResults(data);
    } else if (type === "queryResults") {
      this.handleQueryResults(data);
    } else if (type === "queryError") {
      this.handleQueryError(data);
    } else if (type === "postRetrieveLocalCheck") {
      this.handlePostRetrieveLocalCheck(data);
    } else if (type === "retrieveState") {
      this.handleRetrieveState(data);
    }
  }


  /**
 * Handles retrieve state updates from the extension, toggling the isRetrieving flag
 * to coordinate UI disabled states and progress indicators. [web:31]
 *
 * @param {Object} data Payload containing an isRetrieving boolean field.
 */

  handleRetrieveState(data) {
    if (data && typeof data.isRetrieving === "boolean") {
      this.isRetrieving = data.isRetrieving;
    }
  }


  
/**
 * Handles image resource metadata from the extension, such as a feature logo image
 * used in the Easter egg modal. [web:31]
 *
 * @param {Object} data Payload containing an images map.
 */

  handleImageResources(data) {
    if (data && data?.images?.featureLogo) {
      this.imgFeatureLogo = data.images.featureLogo;
    }
  }

 
  /**
 * Updates metadata rows with LocalFileExists annotations after a retrieve operation
 * and removes successfully retrieved rows from the current selection. [web:31]
 *
 * @param {Object} data Payload containing 'files' and optional 'deletedFiles' arrays.
 */

  handlePostRetrieveLocalCheck(data) {
    // data.files contains annotated records with MemberType, MemberName, LocalFileExists
    const updates = new Map();
    for (const f of data.files) {
      const key = `${f.MemberType}::${f.MemberName}`;
      updates.set(key, f.LocalFileExists);
    }

    // Update existing metadata entries if present
    let changed = false;
    this.metadata = this.metadata.map((row) => {
      const key = `${row.MemberType}::${row.MemberName}`;
      if (updates.has(key)) {
        const exists = updates.get(key);
        changed = true;
        return { ...row, LocalFileIcon: exists === true ? "✅" : "❌" };
      }
      return row;
    });

    // Also unselect any rows that were successfully retrieved (present in data.files)
    try {
      const keysToRemove = new Set();
      for (const f of [...data.files, ...(data.deletedFiles || [])]) {
        const k = `${f.MemberType || f.memberType}::${f.MemberName || f.memberName}`;
        keysToRemove.add(k);
      }

      if (this.selectedRowKeys && this.selectedRowKeys.length > 0) {
        const beforeCount = this.selectedRowKeys.length;
        this.selectedRowKeys = this.selectedRowKeys.filter(
          (k) => !keysToRemove.has(k),
        );
        // Recompute selectedRows based on remaining selectedRowKeys
        this.selectedRows = this.metadata.filter((row) =>
          this.selectedRowKeys.includes(row.uniqueKey),
        );
        if (this.selectedRowKeys.length !== beforeCount) {
          changed = true;
        }
      }
    } catch (e) {
      // non-fatal
    }

    if (changed) {
      // Re-apply client-side filters to refresh the datatable
      this.applyFilters();
      // Re-evaluate floating button visibility since selection/rows might have changed
      setTimeout(() => this.checkRetrieveButtonVisibility(), 0);
    }
  }


  /**
 * Handles results from the listOrgs request and optionally triggers loading
 * of packages for a pre-selected org. [web:16][web:31]
 *
 * @param {Object} data Payload containing 'orgs' and optional 'selectedOrgUsername'.
 */

  handleOrgResults(data) {
    this.isLoadingOrgs = false;
    if (data && data.orgs && Array.isArray(data.orgs)) {
      this.orgs = data.orgs;
      // Set default org if provided or use first available
      if (data.selectedOrgUsername) {
        this.selectedOrg = data.selectedOrgUsername;
        // Trigger package loading for the selected org
        this.isLoadingPackages = true;
        window.sendMessageToVSCode({
          type: "listPackages",
          data: { username: this.selectedOrg },
        });
      }
    }
  }


  /**
 * Handles results from the listPackages request and populates the packageOptions
 * combobox, falling back to default values when none are returned. [web:16]
 *
 * @param {Object} data Payload containing an optional 'packages' array.
 */

  handleListPackagesResults(data) {
    this.isLoadingPackages = false;
    if (data && data.packages && Array.isArray(data.packages)) {
      this.packageOptions = data.packages;
    } else {
      // Fallback to default options
      this.packageOptions = [
        { label: "All", value: "All" },
        { label: "Local", value: "Local" },
      ];
    }
  }


  /**
 * Handles results from the listMetadataTypes request and stores the available
 * metadata types for the current org. [web:16]
 *
 * @param {Object} data Payload containing 'metadataTypes' array.
 */

  handleListMetadataTypesResults(data) {
    if (data && data.metadataTypes && Array.isArray(data.metadataTypes)) {
      this.metadataTypes = data.metadataTypes;
    }
  }


  /**
 * Handles successful query results from the backend by normalizing records
 * into the internal metadata shape and deriving helper fields like ChangeIcon. [web:31]
 *
 * @param {Object} data Payload containing 'records' array.
 */

  handleQueryResults(data) {
    this.isLoading = false;
    if (data && data.records && Array.isArray(data.records)) {
      // Transform records - handle both SourceMember (nested) and Metadata API (flat) formats
      this.metadata = data.records.map((record) => {
        // Use Operation from backend (created/modified/deleted) — guaranteed to be set
        const opVal = (record.Operation || "").toString().toLowerCase();
        // Map operation to colored emoji: created -> green, modified -> yellow, deleted -> red
        let icon = "🟡"; // default = modified
        if (opVal === "created") {
          icon = "🟢";
        } else if (opVal === "deleted") {
          icon = "🔴";
        }

        return {
          MemberName: record.MemberName,
          MemberType: record.MemberType,
          MemberTypeUrl: `${METADATA_DOC_BASE_URL}${record.MemberType}`,
          MemberTypeTitle: `View ${record.MemberType} documentation`,
          MemberNameTitle: `Open metadata for ${record.MemberType} ${record.MemberName}`,
          LastModifiedDate: record.LastModifiedDate,
          // Handle both SourceMember format (LastModifiedBy.Name) and Metadata API format (lastModifiedByName)
          LastModifiedByName:
            record.LastModifiedByName ||
            (record.LastModifiedBy ? record.LastModifiedBy.Name : "") ||
            "",
          uniqueKey: `${record.MemberType}::${record.MemberName}`,
          ChangeIcon: icon,
          // Local file indicator: show  when present; otherwise leave empty
          LocalFileIcon: record.LocalFileExists === true ? "✔️" : "",
        };
      });
      this.applyFilters();
    } else {
      this.metadata = [];
      this.filteredMetadata = [];
    }
  }


  
/**
 * Handles query error messages from the backend by storing the error text
 * and clearing any existing metadata results. [web:31]
 *
 * @param {Object} data Payload containing an optional 'message' field.
 */

  handleQueryError(data) {
    this.isLoading = false;
    this.error =
      data && data.message
        ? data.message
        : "An error occurred while querying metadata";
    this.metadata = [];
    this.filteredMetadata = [];
  }


  /**
 * Handles sort events from the datatable, capturing the sort field and direction
 * and delegating to sortData to reorder filteredMetadata. [web:31]
 *
 * @param {CustomEvent} event lightning-datatable onsort event.
 */

  handleSort(event) {
    const { fieldName, sortDirection } = event.detail;
    this.sortBy = fieldName;
    this.sortDirection = sortDirection;
    this.sortData(fieldName, sortDirection);
  }


  /**
 * Sorts the filteredMetadata array by the given field and direction using a
 * shallow copy to avoid mutating the original array reference. [web:31]
 *
 * @param {string} fieldName Field name to sort by.
 * @param {'asc'|'desc'} direction Sort direction.
 */

  sortData(fieldName, direction) {
    const parseData = JSON.parse(JSON.stringify(this.filteredMetadata));
    const keyValue = (a) => {
      return a[fieldName];
    };
    const isReverse = direction === "asc" ? 1 : -1;
    parseData.sort((x, y) => {
      x = keyValue(x) ? keyValue(x) : "";
      y = keyValue(y) ? keyValue(y) : "";
      return isReverse * ((x > y) - (y > x));
    });
    this.filteredMetadata = parseData;
  }

  /**
 * Shows or hides the floating retrieve button depending on whether the main
 * retrieve button is visible in the viewport and whether there are selected rows. [web:31]
 */

  checkRetrieveButtonVisibility() {
    try {
      const floating = this.template.querySelector(
        '[data-id="retrieve-button-floating"]',
      );
      const mainBtn = this.template.querySelector(
        '[data-id="retrieve-button"]',
      );

      if (!floating) {
        return;
      }

      // If no selected rows, always hide floating button
      if (!this.hasSelectedRows) {
        floating.classList.remove("visible");
        return;
      }

      // If main button is not present in DOM, show floating button
      if (!mainBtn) {
        floating.classList.add("visible");
        return;
      }

      // Check if main button is fully visible in the viewport
      const rect = mainBtn.getBoundingClientRect();
      const viewportHeight =
        window.innerHeight || document.documentElement.clientHeight;
      const isFullyVisible = rect.top >= 0 && rect.bottom <= viewportHeight;

      if (isFullyVisible) {
        floating.classList.remove("visible");
      } else {
        floating.classList.add("visible");
      }
    } catch (e) {
      // In case of any unexpected DOM issues, hide the floating button to be safe
      try {
        const floating = this.template.querySelector(
          '[data-id="retrieve-button-floating"]',
        );
        if (floating) {
          floating.classList.remove("visible");
        }
      } catch (e2) {
        // swallow
      }
    }
  }
}