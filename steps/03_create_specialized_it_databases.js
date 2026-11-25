import notion from "@pipedream/notion"

export default defineComponent({
  name: "Create Database at Workspace Level",
  description: "Create a new database at the workspace level to avoid archived ancestor issues",
  type: "action",
  props: {
    notion,
    title: {
      type: "string",
      label: "Database Title",
      description: "The title of the database",
    },
    properties: {
      type: "string[]",
      label: "Properties",
      description: "Define database properties as JSON objects. Each property should be formatted like: `{\"name\": \"Task Name\", \"type\": \"title\"}` or `{\"name\": \"Status\", \"type\": \"select\", \"select\": {\"options\": [{\"name\": \"To Do\", \"color\": \"red\"}, {\"name\": \"Done\", \"color\": \"green\"}]}}`",
    },
    description: {
      type: "string",
      label: "Database Description",
      description: "Optional description for the database",
      optional: true,
    },
  },
  async run({ $ }) {
    // Parse properties from JSON strings
    const parsedProperties = {};
    
    for (const propString of this.properties) {
      let prop;
      try {
        prop = JSON.parse(propString);
      } catch (error) {
        throw new Error(`Invalid JSON in property: ${propString}. Error: ${error.message}`);
      }
      if (!prop.name || !prop.type) {
        throw new Error(`Each property must have a 'name' and 'type' field: ${propString}`);
      }

      // Notion needs an object keyed by the type, even an empty one (for example "title": {}).
      parsedProperties[prop.name] = {
        type: prop.type,
        [prop.type]: prop[prop.type] ?? {},
      };
    }
    
    // Ensure at least one property exists
    if (Object.keys(parsedProperties).length === 0) {
      throw new Error("At least one property is required");
    }
    
    // Build database object
    const databaseData = {
      parent: {
        type: "workspace",
        workspace: true,
      },
      title: [
        {
          type: "text",
          text: {
            content: this.title,
          },
        },
      ],
      // The Notion app sends Notion-Version 2025-09-03, where a new database's
      // properties go under initial_data_source.
      initial_data_source: {
        properties: parsedProperties,
      },
    };
    
    // Add description if provided
    if (this.description) {
      databaseData.description = [
        {
          type: "text",
          text: {
            content: this.description,
          },
        },
      ];
    }
    
    // Create the database
    const database = await this.notion.createDatabase(databaseData);
    
    $.export("$summary", `Successfully created database "${this.title}" at workspace level`);
    
    return database;
  },
})