/**
 * OpenAPI 3.0 specification for the Enterprise Meeting Search API.
 */

export const openApiSpec = {
  openapi: '3.0.3',
  info: {
    title: 'Enterprise Meeting Search API',
    version: '1.0.0',
    description:
      'Cross-domain meeting search and reconstruction API for Microsoft 365 data sources.',
    contact: {
      name: 'Enterprise Meeting Search',
    },
  },
  servers: [{ url: '/', description: 'Current server' }],
  tags: [
    { name: 'Search', description: 'Cross-domain meeting search endpoints' },
    { name: 'Reconstruction', description: 'Meeting reconstruction from search results' },
    { name: 'Health', description: 'Service health check' },
  ],
  paths: {
    '/health': {
      get: {
        summary: 'Health check',
        tags: ['Health'],
        responses: {
          '200': {
            description: 'Service is healthy',
            content: {
              'application/json': {
                schema: {
                  type: 'object',
                  properties: {
                    status: { type: 'string', example: 'ok' },
                    timestamp: { type: 'string', format: 'date-time' },
                  },
                },
              },
            },
          },
        },
      },
    },
    '/api/v1/search': {
      post: {
        summary: 'Execute a structured cross-domain meeting search',
        tags: ['Search'],
        description:
          'Accepts structured search criteria (keywords, date range, attendees, meeting type) ' +
          'and fans out across all Microsoft 365 data domains in parallel. ' +
          'Returns a searchId that can be polled for results.',
        requestBody: {
          required: true,
          content: {
            'application/json': {
              schema: { $ref: '#/components/schemas/SearchRequest' },
              examples: {
                byKeyword: {
                  summary: 'Search by keyword and date',
                  value: {
                    keywords: ['budget', 'Q4'],
                    dateRange: { start: '2024-01-01T00:00:00Z', end: '2024-03-31T23:59:59Z' },
                  },
                },
                byAttendee: {
                  summary: 'Search by attendee',
                  value: { attendees: ['john.smith@example.com'] },
                },
              },
            },
          },
        },
        responses: {
          '202': {
            description: 'Search accepted and running asynchronously',
            content: {
              'application/json': {
                schema: { $ref: '#/components/schemas/AsyncSearchAccepted' },
              },
            },
          },
          '422': { $ref: '#/components/responses/ValidationError' },
          '429': { $ref: '#/components/responses/RateLimitError' },
          '500': { $ref: '#/components/responses/InternalError' },
        },
      },
    },
    '/api/v1/search/natural': {
      post: {
        summary: 'Natural language meeting search',
        tags: ['Search'],
        description:
          'Accepts a free-text hint (e.g. "mid-January, John Smith was there") and ' +
          'automatically parses it into structured search criteria before executing the search.',
        requestBody: {
          required: true,
          content: {
            'application/json': {
              schema: { $ref: '#/components/schemas/NaturalSearchRequest' },
              examples: {
                midJanuary: {
                  summary: 'Mid January with attendee',
                  value: { hint: 'mid-January, John Smith was there' },
                },
                lastWeekTeams: {
                  summary: 'Last week Teams meeting',
                  value: { hint: 'Teams meeting last week about the product roadmap' },
                },
              },
            },
          },
        },
        responses: {
          '202': {
            description: 'Search accepted and running asynchronously',
            content: {
              'application/json': {
                schema: { $ref: '#/components/schemas/NaturalSearchAccepted' },
              },
            },
          },
          '422': { $ref: '#/components/responses/ValidationError' },
          '429': { $ref: '#/components/responses/RateLimitError' },
          '500': { $ref: '#/components/responses/InternalError' },
        },
      },
    },
    '/api/v1/search/{searchId}/results': {
      get: {
        summary: 'Retrieve async search results',
        tags: ['Search'],
        description: 'Poll this endpoint to retrieve the results of a previously submitted search.',
        parameters: [
          {
            in: 'path',
            name: 'searchId',
            required: true,
            schema: { type: 'string', format: 'uuid' },
            description: 'The search ID returned by POST /api/v1/search or POST /api/v1/search/natural',
          },
        ],
        responses: {
          '200': {
            description: 'Search results (may still be running)',
            content: {
              'application/json': {
                schema: { $ref: '#/components/schemas/SearchResultsResponse' },
              },
            },
          },
          '404': { $ref: '#/components/responses/NotFoundError' },
          '500': { $ref: '#/components/responses/InternalError' },
        },
      },
    },
    '/api/v1/reconstruct/{meetingId}': {
      post: {
        summary: 'Trigger meeting reconstruction',
        tags: ['Reconstruction'],
        description:
          'Reconstructs a structured meeting summary (title, date/time, attendees, topics, ' +
          'decisions, action items) from any previously collected search results referencing ' +
          'the given meeting ID. Run a search first to populate meeting data.',
        parameters: [
          {
            in: 'path',
            name: 'meetingId',
            required: true,
            schema: { type: 'string' },
            description: 'The meeting ID to reconstruct',
          },
        ],
        responses: {
          '200': {
            description: 'Reconstruction completed synchronously',
            content: {
              'application/json': {
                schema: { $ref: '#/components/schemas/ReconstructionResult' },
              },
            },
          },
          '202': {
            description: 'Reconstruction triggered asynchronously',
            content: {
              'application/json': {
                schema: { $ref: '#/components/schemas/ReconstructionAccepted' },
              },
            },
          },
          '404': { $ref: '#/components/responses/NotFoundError' },
          '429': { $ref: '#/components/responses/RateLimitError' },
          '500': { $ref: '#/components/responses/InternalError' },
        },
      },
    },
  },
  components: {
    schemas: {
      DateRange: {
        type: 'object',
        required: ['start', 'end'],
        properties: {
          start: { type: 'string', format: 'date-time', description: 'Inclusive start of the date range' },
          end: { type: 'string', format: 'date-time', description: 'Inclusive end of the date range' },
        },
      },
      SearchRequest: {
        type: 'object',
        description: 'At least one search criterion must be provided.',
        properties: {
          keywords: { type: 'array', items: { type: 'string' }, description: 'Keywords to match across all domains' },
          dateRange: { $ref: '#/components/schemas/DateRange' },
          attendees: { type: 'array', items: { type: 'string', format: 'email' }, description: 'Filter by attendee email addresses' },
          organizer: { type: 'string', format: 'email', description: 'Filter by organizer email address' },
          meetingType: { type: 'array', items: { type: 'string' }, description: 'Filter by meeting type (e.g. Teams, Zoom, InPerson)' },
          domainTimeoutMs: { type: 'integer', minimum: 1, maximum: 30000, description: 'Per-domain timeout in ms (default: 5000)' },
          overallTimeoutMs: { type: 'integer', minimum: 1, maximum: 60000, description: 'Overall search timeout in ms (default: 10000)' },
        },
      },
      NaturalSearchRequest: {
        type: 'object',
        required: ['hint'],
        properties: {
          hint: {
            type: 'string',
            minLength: 3,
            maxLength: 500,
            description: 'Free-text search hint (e.g. "mid-January, John Smith was there")',
            example: 'Teams meeting last week about budget review with John Smith',
          },
        },
      },
      AsyncSearchAccepted: {
        type: 'object',
        properties: {
          searchId: { type: 'string', format: 'uuid' },
          status: { type: 'string', enum: ['running'] },
          message: { type: 'string' },
          resultsUrl: { type: 'string' },
        },
      },
      NaturalSearchAccepted: {
        allOf: [
          { $ref: '#/components/schemas/AsyncSearchAccepted' },
          {
            type: 'object',
            properties: {
              hint: { type: 'string' },
              parsedRequest: { $ref: '#/components/schemas/SearchRequest' },
            },
          },
        ],
      },
      SearchResultsResponse: {
        type: 'object',
        properties: {
          searchId: { type: 'string', format: 'uuid' },
          status: { type: 'string', enum: ['pending', 'running', 'completed', 'failed'] },
          results: {
            type: 'object',
            description: 'The SearchResponse from the orchestrator (present when status is completed)',
          },
          error: { type: 'string', description: 'Error message (present when status is failed)' },
          createdAt: { type: 'string', format: 'date-time' },
          completedAt: { type: 'string', format: 'date-time' },
        },
      },
      ReconstructedElement: {
        type: 'object',
        properties: {
          value: {},
          confidence: { type: 'string', enum: ['high', 'medium', 'low', 'inferred'] },
          source: { type: 'string', enum: ['transcript', 'chat', 'email', 'file', 'calendar'] },
          sourceDetails: { type: 'string' },
        },
      },
      MeetingReconstruction: {
        type: 'object',
        properties: {
          meetingId: { type: 'string' },
          title: { $ref: '#/components/schemas/ReconstructedElement' },
          dateTime: { $ref: '#/components/schemas/ReconstructedElement' },
          attendees: { $ref: '#/components/schemas/ReconstructedElement' },
          topics: { $ref: '#/components/schemas/ReconstructedElement' },
          decisions: { $ref: '#/components/schemas/ReconstructedElement' },
          actionItems: { $ref: '#/components/schemas/ReconstructedElement' },
          overallConfidence: { type: 'string', enum: ['high', 'medium', 'low', 'inferred'] },
          gaps: { type: 'array', items: { type: 'string' } },
          reconstructedAt: { type: 'string', format: 'date-time' },
        },
      },
      ReconstructionResult: {
        type: 'object',
        properties: {
          reconstructionId: { type: 'string', format: 'uuid' },
          meetingId: { type: 'string' },
          status: { type: 'string', enum: ['completed'] },
          reconstruction: { $ref: '#/components/schemas/MeetingReconstruction' },
          completedAt: { type: 'string', format: 'date-time' },
        },
      },
      ReconstructionAccepted: {
        type: 'object',
        properties: {
          reconstructionId: { type: 'string', format: 'uuid' },
          meetingId: { type: 'string' },
          status: { type: 'string', enum: ['running'] },
          message: { type: 'string' },
        },
      },
    },
    responses: {
      ValidationError: {
        description: 'Validation error',
        content: {
          'application/json': {
            schema: {
              type: 'object',
              properties: {
                error: { type: 'string' },
                details: { type: 'array', items: { type: 'object' } },
              },
            },
          },
        },
      },
      NotFoundError: {
        description: 'Resource not found',
        content: {
          'application/json': {
            schema: {
              type: 'object',
              properties: { error: { type: 'string' } },
            },
          },
        },
      },
      RateLimitError: {
        description: 'Rate limit exceeded',
        content: {
          'application/json': {
            schema: {
              type: 'object',
              properties: { error: { type: 'string' } },
            },
          },
        },
      },
      InternalError: {
        description: 'Internal server error',
        content: {
          'application/json': {
            schema: {
              type: 'object',
              properties: { error: { type: 'string' } },
            },
          },
        },
      },
    },
  },
};
