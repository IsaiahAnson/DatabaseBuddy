import notion from "@pipedream/notion"

// The ten categories that have a database prop below. Keep this list in sync with step 01.
const CATEGORIES = [
  "Network Infrastructure",
  "VDI",
  "Applications",
  "Hardware",
  "Security",
  "MDM",
  "Cloud Services",
  "Backup Recovery",
  "User Management",
  "Monitoring",
];

export default defineComponent({
  name: "Populate IT Databases in Notion",
  description: "Create pages in specialized IT databases in Notion with cleaned data from previous steps, organizing content by categories like Network Infrastructure, VDI, Applications, Hardware, Security, etc.",
  type: "action",
  props: {
    notion,
    cleanedData: {
      type: "string[]",
      label: "Cleaned Data",
      description: "Records as JSON strings, in the record shape from the README: id, title, category, content_type, tags, summary, content, source and webUrl",
    },
    networkInfrastructureDb: {
      propDefinition: [
        notion,
        "dataSourceId",
      ],
      label: "Network Infrastructure Database ID",
      description: "The Notion database ID for Network Infrastructure content",
      optional: true,
    },
    vdiDb: {
      propDefinition: [
        notion,
        "dataSourceId",
      ],
      label: "VDI Database ID", 
      description: "The Notion database ID for VDI content",
      optional: true,
    },
    applicationsDb: {
      propDefinition: [
        notion,
        "dataSourceId",
      ],
      label: "Applications Database ID",
      description: "The Notion database ID for Applications content", 
      optional: true,
    },
    hardwareDb: {
      propDefinition: [
        notion,
        "dataSourceId",
      ],
      label: "Hardware Database ID",
      description: "The Notion database ID for Hardware content",
      optional: true,
    },
    securityDb: {
      propDefinition: [
        notion,
        "dataSourceId",
      ],
      label: "Security Database ID",
      description: "The Notion database ID for Security content",
      optional: true,
    },
    mdmDb: {
      propDefinition: [
        notion,
        "dataSourceId",
      ],
      label: "MDM Database ID",
      description: "The Notion database ID for MDM content",
      optional: true,
    },
    cloudServicesDb: {
      propDefinition: [
        notion,
        "dataSourceId",
      ],
      label: "Cloud Services Database ID",
      description: "The Notion database ID for Cloud Services content",
      optional: true,
    },
    backupRecoveryDb: {
      propDefinition: [
        notion,
        "dataSourceId",
      ],
      label: "Backup Recovery Database ID",
      description: "The Notion database ID for Backup Recovery content",
      optional: true,
    },
    userManagementDb: {
      propDefinition: [
        notion,
        "dataSourceId",
      ],
      label: "User Management Database ID",
      description: "The Notion database ID for User Management content",
      optional: true,
    },
    monitoringDb: {
      propDefinition: [
        notion,
        "dataSourceId",
      ],
      label: "Monitoring Database ID",
      description: "The Notion database ID for Monitoring content",
      optional: true,
    },
  },
  methods: {
    getDatabaseMapping() {
      return {
        "Network Infrastructure": this.networkInfrastructureDb,
        "VDI": this.vdiDb,
        "Applications": this.applicationsDb,
        "Hardware": this.hardwareDb,
        "Security": this.securityDb,
        "MDM": this.mdmDb,
        "Cloud Services": this.cloudServicesDb,
        "Backup Recovery": this.backupRecoveryDb,
        "User Management": this.userManagementDb,
        "Monitoring": this.monitoringDb,
      };
    },
    normalizeCategory(category) {
      const categoryMap = {
        "network": "Network Infrastructure",
        "networking": "Network Infrastructure",
        "infrastructure": "Network Infrastructure",
        "vdi": "VDI",
        "virtual desktop": "VDI",
        "app": "Applications",
        "application": "Applications",
        "software": "Applications",
        "hardware": "Hardware",
        "device": "Hardware",
        "equipment": "Hardware",
        "security": "Security",
        "cybersecurity": "Security",
        "mdm": "MDM",
        "mobile device": "MDM",
        "cloud": "Cloud Services",
        "aws": "Cloud Services",
        "azure": "Cloud Services",
        "backup": "Backup Recovery",
        "recovery": "Backup Recovery",
        "disaster recovery": "Backup Recovery",
        "user": "User Management",
        "identity": "User Management",
        "access": "User Management",
        "monitor": "Monitoring",
        "monitoring": "Monitoring",
        "observability": "Monitoring",
        "intune": "MDM",
        "autopilot": "MDM",
        "jamf": "MDM",
        "endpoint": "MDM",
        "entra": "User Management",
        "active directory": "User Management",
        "azure ad": "User Management",
        "avd": "VDI",
        "citrix": "VDI",
        "bitlocker": "Security",
        "defender": "Security",
        "sccm": "Applications",
      };
      
      const value = String(category || "").trim();
      const lowerCategory = value.toLowerCase();

      // 1. A canonical name in any case.
      const canonical = CATEGORIES.find(name => name.toLowerCase() === lowerCategory);
      if (canonical) return canonical;

      // 2. An exact alias.
      if (categoryMap[lowerCategory]) return categoryMap[lowerCategory];

      // 3. An alias as a whole word inside a longer label, longest alias first,
      //    so "Azure AD Connect" maps to User Management, not Cloud Services.
      const aliases = Object.keys(categoryMap).sort((a, b) => b.length - a.length);
      for (const alias of aliases) {
        const pattern = new RegExp(`\\b${alias.replace(/[.*+?^${}()|[\]\\]/g, "\\$&")}\\b`);
        if (pattern.test(lowerCategory)) return categoryMap[alias];
      }

      return value;
    },
    // Notion select option names cannot contain commas and are limited to 100 characters.
    selectName(value) {
      return String(value).replace(/,/g, " ").trim().slice(0, 100);
    },
    // Filter for the "Source ID" property, used to find a page created on an earlier run.
    sourceIdFilter(sourceId) {
      return {
        property: "Source ID",
        rich_text: {
          equals: String(sourceId),
        },
      };
    },
    async findPageBySourceId(dataSourceId, sourceId) {
      const response = await this.notion.queryDataSource(dataSourceId, {
        filter: this.sourceIdFilter(sourceId),
        page_size: 1,
      });
      return response.results?.[0] || null;
    },
    createPageProperties(dataItem) {
      const currentDate = new Date().toISOString();
      const contentType = this.selectName(dataItem.content_type || "");
      
      return {
        "Name": {
          title: [
            {
              text: {
                content: dataItem.title || "Untitled",
              },
            },
          ],
        },
        "Source": {
          rich_text: [
            {
              text: {
                content: dataItem.source || "",
              },
            },
          ],
        },
        "Source ID": {
          rich_text: [
            {
              text: {
                content: String(dataItem.id || ""),
              },
            },
          ],
        },
        "Source URL": {
          url: dataItem.webUrl || null,
        },
        "Category": {
          select: {
            name: dataItem.category || "",
          },
        },
        "Content Type": {
          select: contentType ? { name: contentType } : null,
        },
        "Tags": {
          multi_select: (dataItem.tags || [])
            .map(tag => this.selectName(tag))
            .filter(Boolean)
            .map(tag => ({ name: tag })),
        },
        "Created Date": {
          date: {
            start: currentDate,
          },
        },
        "Last Modified": {
          date: {
            start: currentDate,
          },
        },
        "Content Summary": {
          rich_text: [
            {
              text: {
                content: String(dataItem.summary || "").slice(0, 2000),
              },
            },
          ],
        },
      };
    },
    createPageContent(content) {
      if (!content) return [];
      
      return [
        {
          object: "block",
          type: "paragraph",
          paragraph: {
            rich_text: [
              {
                type: "text",
                text: {
                  content: content.substring(0, 2000), // Limit content length
                },
              },
            ],
          },
        },
      ];
    },
  },
  async run({ $ }) {
    const databaseMapping = this.getDatabaseMapping();
    const results = {
      created: [],
      records: [],
      errors: [],
      summary: {},
    };

    for (const dataString of this.cleanedData || []) {
      try {
        const dataItem = JSON.parse(dataString);
        
        // Normalize and map category to database
        const normalizedCategory = this.normalizeCategory(dataItem.category || "");
        const dataSourceId = databaseMapping[normalizedCategory];
        
        if (!dataSourceId) {
          results.errors.push({
            item: dataItem,
            error: `No database configured for category: ${normalizedCategory}`,
          });
          continue;
        }

        // Create page properties
        const properties = this.createPageProperties({
          ...dataItem,
          category: normalizedCategory,
        });

        // Look for a page made from the same OneDrive item on an earlier run.
        const existingPage = dataItem.id
          ? await this.findPageBySourceId(dataSourceId, dataItem.id)
          : null;

        let page;
        let action;
        if (existingPage) {
          // Update the properties only. Keep the original created date and page body.
          delete properties["Created Date"];
          page = await this.notion.updatePage(existingPage.id, { properties });
          action = "updated";
        } else {
          // Create page content
          const children = this.createPageContent(dataItem.content);

          // Create the page
          page = await this.notion.createPage({
            parent: {
              data_source_id: dataSourceId,
            },
            properties,
            children,
          });
          action = "created";
        }

        results.created.push({
          id: page.id,
          sourceId: dataItem.id || null,
          title: dataItem.title || "Untitled",
          category: normalizedCategory,
          database: dataSourceId,
          action,
        });
        results.records.push({
          ...dataItem,
          category: normalizedCategory,
          pageId: page.id,
        });

        // Update summary counts
        if (!results.summary[normalizedCategory]) {
          results.summary[normalizedCategory] = 0;
        }
        results.summary[normalizedCategory]++;

      } catch (error) {
        results.errors.push({
          item: dataString,
          error: error.message,
        });
      }
    }

    const totalCreated = results.created.filter(p => p.action === "created").length;
    const totalUpdated = results.created.filter(p => p.action === "updated").length;
    const totalErrors = results.errors.length;
    const categoryCount = Object.keys(results.summary).length;

    // Fail the run when records came in but none reached Notion.
    if (results.created.length === 0 && totalErrors > 0) {
      throw new Error(`No pages were written. ${totalErrors} errors. First error: ${results.errors[0].error}`);
    }

    $.export("$summary", `Created ${totalCreated} and updated ${totalUpdated} pages across ${categoryCount} IT categories. ${totalErrors} errors encountered.`);

    return {
      success: totalErrors === 0,
      created_count: totalCreated,
      updated_count: totalUpdated,
      error_count: totalErrors,
      categories_populated: results.summary,
      created_pages: results.created,
      records: results.records,
      errors: results.errors,
    };
  },
})