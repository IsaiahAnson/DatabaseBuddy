import openai from "@pipedream/openai"

export default defineComponent({
  name: "Enhance Records with AI",
  description: "Use OpenAI to enhance individual records with professional titles, detailed summaries, categorization, and suggested tags",
  type: "action",
  props: {
    openai,
    records: {
      type: "string[]",
      label: "Records",
      description: "Array of normalized records to enhance (each record should be a JSON object as a string)"
    },
    model: {
      propDefinition: [
        openai,
        "chatCompletionModelId"
      ],
      default: "gpt-4o-mini"
    },
    includeRelatedDocuments: {
      type: "boolean",
      label: "Include Related Documents Suggestions",
      description: "Whether to generate suggestions for related documents",
      default: false
    },
    customSections: {
      type: "string[]",
      label: "Custom Sections",
      description: "Additional custom sections to generate for each record (e.g., 'key_insights', 'action_items', 'risk_assessment')",
      optional: true
    }
  },
  methods: {
    // Writes the AI output into the shared record shape (title, summary, tags) that step 07 reads.
    // Other AI fields, such as priority_level and custom sections, are kept as they are.
    applyEnhancement(record, enhancement) {
      const { enhanced_title, detailed_summary, suggested_tags, ...extra } = enhancement
      return {
        ...record,
        ...extra,
        // The model's reply must not replace the ids that link the record to OneDrive and Notion.
        id: record.id,
        pageId: record.pageId,
        title: enhanced_title || record.title || record.name,
        summary: detailed_summary || record.summary || "",
        tags: Array.isArray(suggested_tags) && suggested_tags.length > 0 ? suggested_tags : (record.tags || []),
        category: record.category,
        content_type: record.content_type || extra.content_type || "",
        original_title: record.title || record.name,
        enhancement_timestamp: new Date().toISOString(),
        model_used: this.model
      }
    }
  },
  async run({ $ }) {
    const enhancedRecords = []
    let failedCount = 0

    for (const recordStr of this.records) {
      try {
        const record = JSON.parse(recordStr)
        
        // Build the enhancement prompt
        let prompt = `Analyze the following record and enhance it with professional content. Return your response as a JSON object with the following structure:

{
  "enhanced_title": "A professional, descriptive title",
  "detailed_summary": "A comprehensive summary of the content",
  "categories": ["category1", "category2", "category3"],
  "suggested_tags": ["tag1", "tag2", "tag3", "tag4", "tag5"],
  "priority_level": "High|Medium|Low",
  "content_type": "Document|Email|Report|Note|Other"`

        if (this.includeRelatedDocuments) {
          prompt += `,
  "related_documents": ["suggestion1", "suggestion2", "suggestion3"]`
        }

        if (this.customSections && this.customSections.length > 0) {
          for (const section of this.customSections) {
            prompt += `,
  "${section}": "Generated content for ${section}"`
          }
        }

        prompt += `
}

Record to analyze:
${JSON.stringify(record, null, 2)}

Focus on extracting meaningful insights and creating professional, actionable content. Ensure categories are broad but specific enough to be useful for organization. Tags should be relevant keywords that would help with search and filtering.`

        const completion = await this.openai.createChatCompletion({
          $,
          data: {
            model: this.model,
            messages: [
              {
                role: "system",
                content: "You are a professional content analyzer and enhancer. You specialize in creating clear, professional titles, comprehensive summaries, and useful categorization for various types of records and documents. Always respond with valid JSON."
              },
              {
                role: "user",
                content: prompt
              }
            ],
            temperature: 0.7,
            max_tokens: 1000,
            response_format: { type: "json_object" }
          }
        })

        const enhancementText = completion.generated_message?.content
        if (!enhancementText) {
          throw new Error("No enhancement generated")
        }

        let enhancement
        try {
          enhancement = JSON.parse(enhancementText)
        } catch (parseError) {
          // For example a reply cut off at max_tokens. Count it as a failure so step 07
          // leaves the page as it is instead of writing placeholder text.
          throw new Error(`Model reply was not valid JSON: ${parseError.message}`)
        }

        // Merge original record with enhancements
        const enhancedRecord = this.applyEnhancement(record, enhancement)

        enhancedRecords.push(enhancedRecord)

      } catch (error) {
        failedCount++
        $.export("error", `Failed to enhance record: ${error.message}`)
        
        // Add the original record unchanged, with the error, so step 07 leaves the page as it is
        let originalRecord
        try {
          originalRecord = JSON.parse(recordStr)
        } catch {
          originalRecord = { raw_content: recordStr }
        }

        enhancedRecords.push({
          ...originalRecord,
          enhancement_error: error.message,
          enhancement_timestamp: new Date().toISOString()
        })
      }
    }

    // Fail the run when there were records and none could be enhanced.
    if (this.records.length > 0 && failedCount === this.records.length) {
      throw new Error(`Enhancement failed for all ${failedCount} records. Last error: ${enhancedRecords[enhancedRecords.length - 1].enhancement_error}`)
    }

    $.export("$summary", `Enhanced ${this.records.length - failedCount} of ${this.records.length} records using ${this.model}, ${failedCount} failed`)

    return {
      enhanced_records: enhancedRecords,
      total_processed: this.records.length,
      enhancement_settings: {
        model: this.model,
        include_related_documents: this.includeRelatedDocuments,
        custom_sections: this.customSections || []
      }
    }
  }
})