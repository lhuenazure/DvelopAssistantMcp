import express, { Request, Response } from "express";
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StreamableHTTPServerTransport } from "@modelcontextprotocol/sdk/server/streamableHttp.js";
import { z } from "zod";
import { requestContext, requireHeader } from "./requestContext.js";

const server = new McpServer({
  name: "mcp-streamable-http",
  version: "1.0.0",
});

// Helper function to prefix relative links in citations
function prefixLinks(result: any, prefix: string): any {
  const citations = result?.result?.context?.citations;
  if (citations) {
    for (const citation of citations) {
      console.log(`Original link: ${citation?._links?.sourceUri?.href}`);
      const href = citation?._links?.sourceUri?.href;
      if (href) {
        citation._links.sourceUri.href = prefix + href;
        console.log(`Prefixed link: ${citation._links.sourceUri.href}`);
      }
    }
  }
  return result;
}

// Utility function to introduce a delay
const delay = (ms: number) => new Promise((res) => setTimeout(res, ms));

// Function for sending a request to create a prompt
async function createPrompt(body: unknown, authorization: string): Promise<string> {
  const res = await fetch(`https://m365-dev.d-velop.cloud/d42/api/v1/prompts`, {
    method: "POST",
    headers: { Authorization: authorization, "Content-Type": "application/json" },
    body: JSON.stringify(body),
  });
  if (!res.ok) {
    const text = await res.text().catch(() => "");
    throw new Error(`Upstream error ${res.status} ${res.statusText} - ${text}`);
  }
  const data = await res.json();
  if (!data?.id) throw new Error("Prompt-ID fehlt in der Antwort des Servers.");
  return data.id;
}

// Function for polling the prompt until it is completed
async function pollPrompt(promptId: string, authorization: string): Promise<any> {
  const url = `https://m365-dev.d-velop.cloud/d42/api/v1/prompts/${encodeURIComponent(promptId)}`;
  let status = "";
  let result: any = null;
  while (status !== "Completed") {
    await delay(5000); // 5 Sekunden warten
    const res = await fetch(url, { method: "GET", headers: { Authorization: authorization } });
    if (!res.ok) {
      const text = await res.text().catch(() => "");
      throw new Error(`Polling error ${res.status} ${res.statusText} - ${text}`);
    }
    result = await res.json();
    status = result.status;
  }
  return result;
}

// Register tool ask-assistant
server.registerTool(
  "ask-assistant",
  {
    title: "d.velop pilot request",
    description: "Ask the d.velop pilot about quality assurance documents and get a reponse.",
    inputSchema: {
      question: z.string().min(1).describe("The question to ask the d.velop pilot."),
    },
  },
  async ({ question }) => {
    const authorization = requireHeader("authorization");
    const promptBody = {
      prompt: {
        template: question
      },
      context: {
        type: "assistant", assistantId: "54110538-f15b-4ea2-a88d-264ebe19f790"
      }
    };
    const promptId = await createPrompt(promptBody, authorization);
    const result = await pollPrompt(promptId, authorization);
    const updatedResult = prefixLinks(result, "https://m365-dev.d-velop.cloud");
    console.log("Old Result:", result);
    console.log("Updated Result:", updatedResult);
    return updatedResult;
  }
);

// Output-Shape für list-users
const listUsersOutputShape = {
  totalResults: z.number().describe("Total number of users."),
  itemsPerPage: z.number().describe("Number of items per page."),
  startIndex: z.number().describe("Start index of the current page."),
  resources: z.array(
    z.object({
      id: z.string().describe("The unique user ID."),
      userName: z.string().describe("The username."),
      displayName: z.string().describe("The display name of the user."),
      preferredLanguage: z.string().optional().describe("The user's preferred language."),
      emails: z.array(
        z.object({
          value: z.string().email().describe("The user's email address."),
        })
      ).describe("List of email addresses."),
      photos: z.array(
        z.object({
          value: z.string().describe("The photo URL."),
          type: z.string().describe("The type of photo."),
        })
      ).optional().describe("List of photos."),
      title: z.string().optional().describe("Job title of the user."),
      department: z.string().optional().describe("Department of the user."),
    })
  ).describe("List of user resources."),
} satisfies z.ZodRawShape;

// Zod-Objekt für Validierung
const ListUsersSchema = z.object(listUsersOutputShape);

// Register tool list-users
server.registerTool(
  "list-users",
  {
    title: "List Users",
    description: "Retrieve a list of all users from the SCIM endpoint.",
    inputSchema: {}, // keine Eingaben
    outputSchema: listUsersOutputShape,
  },
  async () => {
    const authorization = requireHeader("authorization");
    const res = await fetch(`https://m365-dev.d-velop.cloud/identityprovider/scim/users`, {
      method: "GET",
      headers: { Authorization: authorization },
    });

    if (!res.ok) {
      const text = await res.text().catch(() => "");
      throw new Error(`Upstream error ${res.status} ${res.statusText} - ${text}`);
    }

    const data = await res.json();

    // Validierung gegen das Schema
    const validated = ListUsersSchema.parse(data);

    // MCP-konforme Antwort zurückgeben
    return {
      content: [
        {
          type: "text",
          text: `Retrieved ${validated.totalResults} users.`,
        },
      ],
      structuredContent: validated,
    };
  }
);

// Input schema for the tool
const createTaskInputShape = {
  subject: z.string().min(1).describe("The subject/title of the task."),
  description: z.string().optional().describe("A descriptive text of the task."),
  assignees: z.array(z.string().min(1)).min(1).describe("List of user IDs or group IDs to assign the task to. Must be retrieved via the list-users tool."),
  dueDate: z.string().regex(
    /^\d{4}-\d{2}-\d{2}(T\d{2}:\d{2}:\d{2}(Z|([+-]\d{2}:\d{2})))?$/,
    "Must be RFC3339 format (e.g., 2025-10-16 or 2025-10-16T00:00:00Z)"
  ).describe("Due date in RFC3339 format."),
} satisfies z.ZodRawShape;

const CreateTaskInputSchema = z.object(createTaskInputShape);

// Register tool create-task
server.registerTool(
  "create-task",
  {
    title: "Create Task",
    description: "Creates a new task in the system with subject, description, assignees, correlation key, and due date.",
    inputSchema: createTaskInputShape,
    // No outputSchema because the API returns an empty body
  },
  async ({ subject, description, assignees, dueDate }) => {
    const authorization = requireHeader("authorization");

    // Generate a unique correlation key
    const correlationKey = `task-${Date.now()}-${Math.floor(Math.random() * 1000)}`;

    const body = {
      subject,
      description,
      assignees,
      correlationKey,
      dueDate,
    };

    const res = await fetch(`https://m365-dev.d-velop.cloud/task/tasks`, {
      method: "POST",
      headers: {
        Authorization: authorization,
        "Content-Type": "application/json",
      },
      body: JSON.stringify(body),
    });

    if (res.status === 201) {
      return {
        content: [
          {
            type: "text",
            text: `✅ Task "${subject}" created successfully for ${assignees.length} assignee(s).`,
          },
        ],
      };
    } else {
      const errorText = await res.text().catch(() => "");
      throw new Error(`Task creation failed: ${res.status} ${res.statusText} - ${errorText}`);
    }
  }
);


// Output shape for list-tasks
const listTasksOutputShape = {
  _embedded: z.object({
    tasks: z.array(
      z.object({
        subject: z.string().describe("The subject or title of the task."),
        description: z.string().optional().describe("Detailed description of the task."),
        assignedUsers: z.array(z.string()).describe("Array of assigned user IDs."),
        assignedGroups: z.array(z.string()).optional().describe("Array of assigned group IDs."),
        senderLabel: z.string().describe("Display name of the sender."),
        sender: z.string().describe("Unique sender ID."),
        receiveDate: z.string().describe("Date when the task was received."),
        dueDate: z.string().describe("Deadline for the task."),
        priority: z.number().describe("Priority level of the task."),
        id: z.string().describe("Unique identifier for the task."),
        completed: z.boolean().describe("Indicates whether the task is completed."),
        editor: z.string().optional().describe("Editor ID if applicable."),
        editorLabel: z.string().optional().describe("Display name of the editor."),
        correlationKey: z.string().optional().describe("Correlation key for idempotency."),
        readByCurrentUser: z.boolean().describe("Indicates if the current user has read the task."),
        orderValue: z.number().describe("Numeric value for sorting tasks."),
        retentionTime: z.string().optional().describe("Retention period (e.g., P30D)."),
        lockHolder: z.string().optional().describe("ID of the user holding the lock."),
        dmsReferences: z.array(z.string()).optional().describe("References to related DMS documents."),
        actionScopes: z.record(z.array(z.string())).describe("Available actions and their scopes."),
        undelivered: z.boolean().describe("Indicates if the task was undelivered."),
        _links: z.record(z.any()).describe("Hyperlinks related to the task."),
      })
    ),
  }),
  _links: z.record(z.any()).describe("Links for the search result."),
} satisfies z.ZodRawShape;

const ListTasksSchema = z.object(listTasksOutputShape);

// Register tool list-tasks
server.registerTool(
  "list-tasks",
  {
    title: "List Tasks",
    description: "Retrieve all tasks sorted by received date in ascending order.",
    inputSchema: {}, // No input required, body is fixed
    outputSchema: listTasksOutputShape,
  },
  async () => {
    const authorization = requireHeader("authorization");

    const body = {
      orderBy: "received",
      completed: false,
    };

    const res = await fetch(`https://m365-dev.d-velop.cloud/task/tasks/search`, {
      method: "POST",
      headers: {
        Authorization: authorization,
        "Content-Type": "application/json",
      },
      body: JSON.stringify(body),
    });

    if (!res.ok) {
      const errorText = await res.text().catch(() => "");
      throw new Error(`Failed to retrieve tasks: ${res.status} ${res.statusText} - ${errorText}`);
    }

    const data = await res.json();

    // Validate against schema
    const validated = ListTasksSchema.parse(data);

    // Return MCP-compliant response
    return {
      content: [
        {
          type: "text",
          text: `✅ Retrieved ${validated._embedded.tasks.length} task(s). First task: "${validated._embedded.tasks[0]?.subject ?? "none"}".`,
        },
      ],
      structuredContent: validated,
    };
  }
);

const app = express();
app.use(express.json());

const transport: StreamableHTTPServerTransport =
  new StreamableHTTPServerTransport({
    sessionIdGenerator: undefined, // set to undefined for stateless servers
  });

// Setup routes for the server
const setupServer = async () => {
  await server.connect(transport);
};

app.post("/mcp", async (req: Request, res: Response) => {
  await requestContext.run({ headers: req.headers }, async () => {
    try {
      await transport.handleRequest(req, res, req.body);
    } catch (error) {
      console.error("Error handling MCP request:", error);
      if (!res.headersSent) {
        res.status(500).json({
          jsonrpc: "2.0",
          error: { code: -32603, message: "Internal server error" },
          id: null,
        });
      }
    }
  });
});

app.get("/mcp", async (req: Request, res: Response) => {
  console.log("Received GET MCP request");
  res.writeHead(405).end(
    JSON.stringify({
      jsonrpc: "2.0",
      error: {
        code: -32000,
        message: "Method not allowed.",
      },
      id: null,
    })
  );
});

app.delete("/mcp", async (req: Request, res: Response) => {
  console.log("Received DELETE MCP request");
  res.writeHead(405).end(
    JSON.stringify({
      jsonrpc: "2.0",
      error: {
        code: -32000,
        message: "Method not allowed.",
      },
      id: null,
    })
  );
});

// Start the server
const PORT = process.env.PORT || 3000;
setupServer()
  .then(() => {
    app.listen(PORT, () => {
      console.log(`MCP Streamable HTTP Server listening on port ${PORT}`);
    });
  })
  .catch((error) => {
    console.error("Failed to set up the server:", error);
    process.exit(1);
  });
