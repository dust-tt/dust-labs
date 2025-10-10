function convertMarkdownToHtml(markdown, conversationData) {
  // Check if markdown is null or not a string
  if (typeof markdown !== "string") {
    console.error("Invalid markdown input:", markdown);
    return ""; // Return an empty string or handle as needed
  }

  // Clean up leading newlines and whitespace
  markdown = markdown.replace(/^\s*\n+/, '').trim();

  // Store conversation data
  currentConversation = conversationData;

  // Keep track of citation count
  let citationCount = 0;
  const citations = {};

  // First pass: collect all citations and assign numbers
  markdown.replace(/:cite\[([^\]]+)\]/g, (match, reference) => {
    if (!citations[reference]) {
      citationCount++;
      citations[reference] = citationCount;
    }
  });

  return (
    markdown
      .replace(/^### (.*$)/gim, "<h3>$1</h3>")
      .replace(/^## (.*$)/gim, "<h2>$1</h2>")
      .replace(/^# (.*$)/gim, "<h1>$1</h1>")
      .replace(/\*\*(.*)\*\*/gim, "<strong>$1</strong>")
      .replace(/\*(.*)\*/gim, "<em>$1</em>")
      .replace(/!\[(.*?)\]\((.*?)\)/gim, "<img alt='$1' src='$2' />")
      .replace(/\[(.*?)\]\((.*?)\)/gim, "<a href='$2' target='_blank'>$1</a>")
      .replace(/`(.*?)`/gim, "<code>$1</code>")
      .replace(
        /```([\s\S]*?)```/g,
        (match, p1) => "<pre><code>" + p1.trim() + "</code></pre>"
      )
      .replace(/(?:\r\n|\r|\n)/g, "<br>")
      // Replace citations with numbered badges that are links
      .replace(/:cite\[([^\]]+)\]/g, (match, reference) => {
        const sourceUrl = getSourceUrlFromReference(reference);
        return `<a href="${sourceUrl}" target="_blank" class="citation-badge">[${citations[reference]}]</a>`;
      })
  );
}

// Store conversation data at module level
let currentConversation = null;

// Helper function to get source URL from the conversation data
function getSourceUrlFromReference(reference) {
  if (!currentConversation) {
    console.warn("No conversation data available");
    return "#";
  }

  try {
    const document =
      currentConversation.content[1][0].actions[0].documents.find(
        (doc) => doc.reference === reference
      );
    return document ? document.sourceUrl : "#";
  } catch (error) {
    console.error("Error finding source URL for reference:", error);
    return "#";
  }
}

(async function () {
  const client = ZAFClient.init();
  const isProd = true;
  window.client = client;
  window.useAnswer = useAnswer;
  window.copyMessage = copyMessage;

  let defaultAssistantIds;

  // Helper function to get Dust base URL based on region
  function getDustBaseUrl(metadata) {
    const region = metadata.settings.region;

    if (region && region.toLowerCase() === "eu") {
      return "https://eu.dust.tt";
    }
    return "https://dust.tt";
  }

  try {
    await client.on("app.registered");
    const metadata = await client.metadata();
    defaultAssistantIds = metadata.settings.default_assistant_ids;

    if (
      defaultAssistantIds &&
      defaultAssistantIds.length > 0 &&
      defaultAssistantIds != "undefined"
    ) {
      const assistantIds = defaultAssistantIds
        .split(/[,\s]+/) // Split by comma, space, or newline
        .filter((id) => id.trim() !== "") // Remove empty entries
        .map((id) => id.trim()); // Trim each ID

      await loadAssistants(assistantIds);
      showAssistantSelect();
      restoreSelectedAssistant();
    } else {
      await loadAssistants();
      showAssistantSelect();
      restoreSelectedAssistant();
    }

    hideLoadingSpinner();
  } catch (error) {
    hideLoadingSpinner();
    showErrorMessage(
      error.message || "Failed to load assistants. Please try again later."
    );
  }

  await client.invoke("resize", { width: "100%", height: "400px" });

  const sendToDustButton = document.getElementById("sendToDustButton");
  const userInput = document.getElementById("userInput");

  sendToDustButton.addEventListener("click", handleSubmit);

  userInput.addEventListener("keydown", (event) => {
    if (event.key === "Enter" && !event.shiftKey) {
      event.preventDefault();
      handleSubmit();
    }
  });

  userInput.addEventListener("input", autoResize);

  function autoResize() {
    this.style.height = "auto";
    this.style.height = this.scrollHeight + "px";
  }

  async function handleSubmit() {
    try {
      const data = await client.get("ticket");
      sendTicketToDust(data.ticket);
    } catch (error) {
      console.error("Error getting ticket data:", error);
    }
  }

  async function checkUserValidity(dustWorkspaceId, dustApiKey, userEmail) {
    const metadata = await client.metadata();
    const baseUrl = getDustBaseUrl(metadata);

    const validationUrl = `${baseUrl}/api/v1/w/${dustWorkspaceId}/members/validate`;
    const options = {
      url: validationUrl,
      type: "POST",
      headers: {
        Authorization: `Bearer ${dustApiKey}`,
      },
      secure: isProd,
      data: { email: userEmail },
    };

    try {
      const response = await client.request(options);
      if (response && response.valid) {
        return true;
      } else {
        return false;
      }
    } catch (error) {
      console.error("Error validating user:", error);
      return false;
    }
  }

  async function loadAssistants(allowedAssistantIds = null) {
    const metadata = await client.metadata();
    const dustApiKey = isProd
      ? "{{setting.dust_api_key}}"
      : `${metadata.settings.dust_api_key}`;
    const dustWorkspaceId = isProd
      ? "{{setting.dust_workspace_id}}"
      : `${metadata.settings.dust_workspace_id}`;

    const userData = await client.get("currentUser");
    const userEmail = userData.currentUser.email;

    const isValid = await checkUserValidity(
      dustWorkspaceId,
      dustApiKey,
      userEmail
    );
    if (!isValid) {
      throw new Error(
        "You need a Dust.tt account to use this app. Please contact your administrator to enable access to Dust"
      );
    }

    const authorization = `Bearer ${dustApiKey}`;
    const baseUrl = getDustBaseUrl(metadata);
    const assistantsApiUrl = `${baseUrl}/api/v1/w/${dustWorkspaceId}/assistant/agent_configurations`;

    const options = {
      url: assistantsApiUrl,
      type: "GET",
      headers: {
        Authorization: authorization,
      },
      secure: isProd,
    };

    const response = await client.request(options);
    if (
      response &&
      response.agentConfigurations &&
      Array.isArray(response.agentConfigurations)
    ) {
      let assistants = response.agentConfigurations;

      if (allowedAssistantIds && allowedAssistantIds.length > 0) {
        assistants = assistants.filter((assistant) =>
          allowedAssistantIds.includes(assistant.sId)
        );
      }

      if (assistants.length === 0) {
        throw new Error("No assistants found");
      }

      const selectElement = document.getElementById("assistantSelect");
      const selectWrapper = document.getElementById("assistantSelectWrapper");
      const inputWrapper = document.getElementById("inputWrapper");

      // Clear existing options
      selectElement.innerHTML = "";

      assistants.forEach((assistant) => {
        if (assistant && assistant.sId && assistant.name) {
          const option = new Option(`@${assistant.name}`, assistant.sId);
          // Store the description as a data attribute
          if (assistant.description) {
            option.dataset.description = assistant.description;
          }
          selectElement.appendChild(option);
        }
      });

      $(selectElement)
        .select2({
          placeholder: "Select an assistant",
          allowClear: true,
          templateResult: formatAssistant,
          templateSelection: formatAssistantSelection,
        })
        .on("change", (e) => {
          localStorage.setItem("selectedAssistant", e.target.value);
          localStorage.setItem(
            "selectedAssistantName",
            e.target.options[e.target.selectedIndex].text
          );
        })
        .on("select2:open", () => {
          setTimeout(() => {
            document.querySelector(".select2-search__field").focus();
          }, 0);
        });

      function formatAssistant(assistant) {
        if (!assistant.id) {
          return assistant.text;
        }

        const option = assistant.element;
        const description = option.dataset.description;

        const createAssistantHTML = (name, description = "") => {
          return $(`
            <div class="assistant-option">
              <div class="assistant-name">${name}</div>
              ${
                description
                  ? `<div class="assistant-description">${description}</div>`
                  : ""
              }
            </div>
          `);
        };

        if (!description) {
          return createAssistantHTML(assistant.text);
        }

        const truncatedDescription =
          description.length > 100
            ? `${description.substring(0, 100)}...`
            : description;

        return createAssistantHTML(assistant.text, truncatedDescription);
      }

      function formatAssistantSelection(assistant) {
        if (!assistant.id) {
          return assistant.text;
        }
        return assistant.text;
      }

      selectWrapper.style.display = "block";

      if (assistants.length === 1) {
        $(selectElement).val(assistants[0].sId).trigger("change");
      }

      inputWrapper.style.display = "block";
    } else {
      throw new Error("Unexpected API response format");
    }
  }

  function hideLoadingSpinner() {
    document.getElementById("loadingSpinner").style.display = "none";
  }

  function showAssistantSelect() {
    document.getElementById("assistantSelect").style.display = "block";
  }

  function showErrorMessage(message) {
    const errorElement = document.createElement("div");
    errorElement.textContent = message;
    errorElement.style.color = "grey";
    errorElement.style.textAlign = "center";
    document.getElementById("assistantSelectWrapper").appendChild(errorElement);
  }

  function restoreSelectedAssistant() {
    const savedAssistant = localStorage.getItem("selectedAssistant");
    if (savedAssistant) {
      const selectElement = document.getElementById("assistantSelect");
      if (selectElement.style.display !== "none") {
        $(selectElement).val(savedAssistant).trigger("change");
      }
    }
  }

  async function pollConversationEvents(conversationId, uniqueId, dustWorkspaceId, authorization, baseUrl, metadata) {
    const maxPollingTime = 3 * 60 * 1000; // 3 minutes in milliseconds
    const pollInterval = 1000; // 1 second for faster updates
    const startTime = Date.now();
    
    let lastEventIndex = -1;
    let hasContent = false;
    let currentContent = '';
    let isCompleted = false;
    
    const assistantMessageElement = document.getElementById(`assistant-${uniqueId}`);
    
    while (Date.now() - startTime < maxPollingTime && !isCompleted) {
      try {
        const eventsUrl = `${baseUrl}/api/v1/w/${dustWorkspaceId}/assistant/conversations/${conversationId}`;
        const eventsOptions = {
          url: eventsUrl,
          type: 'GET',
          headers: {
            Authorization: authorization,
          },
          secure: isProd,
        };
        
        const eventsResponse = await client.request(eventsOptions);
        
        if (eventsResponse && eventsResponse.conversation && eventsResponse.conversation.content) {
          const messages = eventsResponse.conversation.content;
          
          // Look for assistant messages (index 1 and beyond)
          for (let i = 1; i < messages.length; i++) {
            const messageGroup = messages[i];
            if (Array.isArray(messageGroup) && messageGroup.length > 0) {
              const message = messageGroup[0];
              
              if (message.type === 'agent_message') {
                hasContent = true;
                const agentName = message.configuration?.name || 'Assistant';
                
                // Debug: log the entire message structure
                console.log('Agent message received:', JSON.stringify(message, null, 2));
                console.log('Message status:', message.status);
                
                // Extract chain of thought if available
                const chainOfThought = message.chainOfThought || '';
                
                // Function to escape HTML and preserve line breaks
                function formatChainOfThought(text) {
                  if (!text || !text.trim()) return '';
                  
                  // Escape HTML special characters
                  const escaped = text
                    .replace(/&/g, '&amp;')
                    .replace(/</g, '&lt;')
                    .replace(/>/g, '&gt;')
                    .replace(/"/g, '&quot;')
                    .replace(/'/g, '&#39;');
                  
                  return `<div class="chain-of-thought" id="chain-of-thought-${uniqueId}">${escaped}</div>`;
                }
                
                // Function to update chain of thought with smooth transition
                function updateChainOfThought(text, elementId) {
                  const existingChainOfThought = document.getElementById(`chain-of-thought-${elementId}`);
                  
                  if (!text || !text.trim()) {
                    if (existingChainOfThought) {
                      existingChainOfThought.classList.add('fade-out');
                      setTimeout(() => {
                        if (existingChainOfThought.parentNode) {
                          existingChainOfThought.remove();
                        }
                      }, 200);
                    }
                    return '';
                  }
                  
                  const escaped = text
                    .replace(/&/g, '&amp;')
                    .replace(/</g, '&lt;')
                    .replace(/>/g, '&gt;')
                    .replace(/"/g, '&quot;')
                    .replace(/'/g, '&#39;');
                  
                  if (existingChainOfThought) {
                    // Update existing content
                    existingChainOfThought.innerHTML = escaped;
                    if (!existingChainOfThought.classList.contains('visible')) {
                      setTimeout(() => existingChainOfThought.classList.add('visible'), 10);
                    }
                    return ''; // Return empty since element already exists in DOM
                  } else {
                    // Create new element and return HTML to insert
                    return `<div class="chain-of-thought" id="chain-of-thought-${elementId}">${escaped}</div>`;
                  }
                }
                
                const chainOfThoughtHtml = formatChainOfThought(chainOfThought);
                
                // Debug logging for chain of thought
                if (chainOfThought && chainOfThought.trim()) {
                  console.log('Chain of thought found:', chainOfThought);
                  console.log('Chain of thought HTML:', chainOfThoughtHtml);
                } else {
                  console.log('No chain of thought found or empty');
                }
                
                // Test: force show chain of thought for debugging (remove this after testing)
                // const testChainOfThought = message.chainOfThought ? formatChainOfThought('• Testing chain of thought display\n• This should appear immediately') : '';
                // if (message.chainOfThought) {
                //   console.log('FORCED TEST: Displaying test chain of thought');
                // }
                
                if (message.status === 'succeeded') {
                  // Message is complete - hide chain of thought and show final answer
                  isCompleted = true;
                  const htmlAnswer = convertMarkdownToHtml(message.content, eventsResponse.conversation);
                  if (assistantMessageElement) {
                    // Hide any existing chain of thought
                    const existingChainOfThought = document.getElementById(`chain-of-thought-${uniqueId}`);
                    if (existingChainOfThought) {
                      existingChainOfThought.classList.add('fade-out');
                      setTimeout(() => {
                        if (existingChainOfThought.parentNode) {
                          existingChainOfThought.remove();
                        }
                      }, 200);
                    }

                    setTimeout(() => {
                      assistantMessageElement.innerHTML = `
                        <div class="message-header">
                          <h4>@${agentName}:</h4>
                          <button class="copy-button" onclick="copyMessage(this)" title="Copy message">
                            <svg viewBox="0 0 24 24">
                              <path d="M16 1H4c-1.1 0-2 .9-2 2v14h2V3h12V1zm3 4H8c-1.1 0-2 .9-2 2v14c0 1.1.9 2 2 2h11c1.1 0 2-.9 2-2V7c0-1.1-.9-2-2-2zm0 16H8V7h11v14z"/>
                            </svg>
                          </button>
                        </div>
                        <pre class=\"markdown-content\" data-markdown="${message.content.replace(/"/g, '&quot;')}">${htmlAnswer}</pre>
                        <button class=\"use-button public-reply\" onclick=\"useAnswer(this, 'public')\">Use as public reply</button>
                        <button class=\"use-button private-note\" onclick=\"useAnswer(this, 'private')\">Use as internal note</button>
                      `;
                    }, 200);
                  }
                } else if (message.status === 'created' || (message.status === 'running' && !message.content)) {
                  // Message is just created or running but no content yet
                  if (assistantMessageElement) {
                    let chainOfThoughtElement = document.getElementById(`chain-of-thought-${uniqueId}`);
                    
                    // Handle chain of thought display
                    if (chainOfThought && chainOfThought.trim()) {
                      const escaped = chainOfThought
                        .replace(/&/g, '&amp;')
                        .replace(/</g, '&lt;')
                        .replace(/>/g, '&gt;')
                        .replace(/"/g, '&quot;')
                        .replace(/'/g, '&#39;');
                      
                      if (!chainOfThoughtElement) {
                        // Create chain of thought element
                        assistantMessageElement.innerHTML = `
                          <h4>@${agentName}:</h4>
                          <div class="chain-of-thought" id="chain-of-thought-${uniqueId}">${escaped}</div>
                          <div class=\"spinner\"></div>
                          <div class=\"generating-status\">Generating response...</div>
                        `;
                        
                        // Trigger animation after DOM update
                        setTimeout(() => {
                          const element = document.getElementById(`chain-of-thought-${uniqueId}`);
                          if (element) {
                            element.classList.add('visible');
                          }
                        }, 50);
                      } else {
                        // Update existing chain of thought content
                        chainOfThoughtElement.innerHTML = escaped;
                        // Auto-scroll to bottom of chain of thought
                        chainOfThoughtElement.scrollTop = chainOfThoughtElement.scrollHeight;
                      }
                    } else {
                      // No chain of thought - just show spinner
                      if (!chainOfThoughtElement) {
                        assistantMessageElement.innerHTML = `
                          <h4>@${agentName}:</h4>
                          <div class=\"spinner\"></div>
                          <div class=\"generating-status\">Generating response...</div>
                        `;
                      }
                    }
                  }
                } else if (message.status === 'running' && message.content) {
                  // Message is still generating but has partial content
                  const htmlAnswer = convertMarkdownToHtml(message.content, eventsResponse.conversation);
                  if (assistantMessageElement) {
                    let chainOfThoughtElement = document.getElementById(`chain-of-thought-${uniqueId}`);
                    
                    // Handle chain of thought display
                    if (chainOfThought && chainOfThought.trim()) {
                      const escaped = chainOfThought
                        .replace(/&/g, '&amp;')
                        .replace(/</g, '&lt;')
                        .replace(/>/g, '&gt;')
                        .replace(/"/g, '&quot;')
                        .replace(/'/g, '&#39;');
                      
                      if (!chainOfThoughtElement) {
                        // Create chain of thought element
                        assistantMessageElement.innerHTML = `
                          <h4>@${agentName}:</h4>
                          <div class="chain-of-thought" id="chain-of-thought-${uniqueId}">${escaped}</div>
                          <pre class=\"markdown-content\">${htmlAnswer}</pre>
                          <div class=\"spinner\"></div>
                          <div class=\"generating-status\">Generating response...</div>
                        `;
                        
                        // Trigger animation after DOM update
                        setTimeout(() => {
                          const element = document.getElementById(`chain-of-thought-${uniqueId}`);
                          if (element) {
                            element.classList.add('visible');
                          }
                        }, 50);
                      } else {
                        // Update existing chain of thought content
                        chainOfThoughtElement.innerHTML = escaped;
                        // Auto-scroll to bottom of chain of thought
                        chainOfThoughtElement.scrollTop = chainOfThoughtElement.scrollHeight;
                      }
                    } else {
                      // No chain of thought - just show content and spinner
                      if (!chainOfThoughtElement) {
                        assistantMessageElement.innerHTML = `
                          <h4>@${agentName}:</h4>
                          <pre class=\"markdown-content\">${htmlAnswer}</pre>
                          <div class=\"spinner\"></div>
                          <div class=\"generating-status\">Generating response...</div>
                        `;
                      }
                    }
                  }
                } else if (message.status === 'errored') {
                  // Message failed
                  isCompleted = true;
                  if (assistantMessageElement) {
                    assistantMessageElement.innerHTML = `
                      <h4>Error:</h4>
                      <pre>Failed to generate response. Please try again.</pre>
                    `;
                  }
                }
                break; // Only process the first agent message
              }
            }
          }
        }
        
        document.getElementById('dustResponse').scrollTop = document.getElementById('dustResponse').scrollHeight;
        
        if (!isCompleted) {
          await new Promise(resolve => setTimeout(resolve, pollInterval));
        }
        
      } catch (error) {
        console.error('Error polling conversation events:', error);
        await new Promise(resolve => setTimeout(resolve, pollInterval));
      }
    }
    
    // If we timed out, show error message
    if (!isCompleted && Date.now() - startTime >= maxPollingTime) {
      if (assistantMessageElement) {
        assistantMessageElement.innerHTML = `
          <h4>Error:</h4>
          <pre>Request timed out after 3 minutes. The assistant may still be processing your request.</pre>
        `;
      }
    }
    
    await client.invoke("resize", { width: "100%", height: "600px" });
  }

  async function sendTicketToDust(ticket) {
    const dustResponse = document.getElementById("dustResponse");
    const userInput = document.getElementById("userInput");
    let uniqueId;

    try {
      const metadata = await client.metadata();
      const dustApiKey = isProd
        ? "{{setting.dust_api_key}}"
        : `${metadata.settings.dust_api_key}`;
      const dustWorkspaceId = isProd
        ? "{{setting.dust_workspace_id}}"
        : `${metadata.settings.dust_workspace_id}`;
      const hideCustomerInformation =
        metadata.settings.hide_customer_information;
      const baseUrl = getDustBaseUrl(metadata);
      const dustApiUrl = `${baseUrl}/api/v1/w/${dustWorkspaceId}/assistant/conversations`;
      const authorization = `Bearer ${dustApiKey}`;

      let selectedAssistantId, selectedAssistantName;
      const selectElement = document.getElementById("assistantSelect");

      if (selectElement.style.display === "none") {
        selectedAssistantId = localStorage.getItem("selectedAssistant");
        selectedAssistantName = localStorage.getItem("selectedAssistantName");
      } else {
        selectedAssistantId = selectElement.value;
        selectedAssistantName =
          selectElement.options[selectElement.selectedIndex].text;
      }

      if (!selectedAssistantId || !selectedAssistantName) {
        throw new Error(
          "No assistant selected. Please select an assistant before sending a message."
        );
      }

      const userInputValue = userInput.value;

      const userData = await client.get("currentUser");
      const userFullName = userData.currentUser.name;

      const data = await client.get("ticket");

      uniqueId = generateUniqueId();

      const ticketInfo = {
        id: ticket.id || "Unknown",
        subject: ticket.subject || "No subject",
        description: ticket.description || "No description",
        status: ticket.status || "Unknown",
        priority: ticket.priority || "Not set",
        type: ticket.type || "Not specified",
        tags: Array.isArray(ticket.tags) ? ticket.tags.join(", ") : "No tags",
        createdAt: ticket.createdAt || "Unknown",
        updatedAt: ticket.updatedAt || "Unknown",
        assignee:
          (ticket.assignee &&
            ticket.assignee.user &&
            ticket.assignee.user.name) ||
          "Unassigned",
        assignee_email:
          (ticket.assignee &&
            ticket.assignee.user &&
            ticket.assignee.user.email) ||
          "Unassigned",
        group: (ticket.group && ticket.group.name) || "No group",
        organization:
          (ticket.organization && ticket.organization.name) ||
          "No organization",
        customerName: hideCustomerInformation ? "[Redacted]" : "Unknown",
        customerEmail: hideCustomerInformation ? "[Redacted]" : "Unknown",
      };

      if (
        data &&
        data.ticket &&
        data.ticket.requester &&
        !hideCustomerInformation
      ) {
        ticketInfo.customerName = data.ticket.requester.name || "Unknown";
        ticketInfo.customerEmail = data.ticket.requester.email || "Unknown";
      }

      dustResponse.innerHTML += `
      <div class="user-message" id="user-${uniqueId}">
        <strong>${userFullName}:</strong>
        <pre>${userInputValue}</pre>
      </div>
    `;

      dustResponse.innerHTML += `
      <div id="${"assistant-" + uniqueId}" class="assistant-message">
        <h4>${selectedAssistantName}:</h4>
        <div class="spinner"></div>
      </div>
    `;

      dustResponse.scrollTop = dustResponse.scrollHeight;

      userInput.value = "";

      const commentsResponse = await client.request(
        `/api/v2/tickets/${ticket.id}/comments.json`
      );
      const comments = commentsResponse.comments;

      const userIds = [
        ...new Set(
          comments
            .map((comment) => comment.author_id)
            .filter((id) => id && id > 0)
        ),
      ];

      const userResponses = await Promise.all(
        userIds.map((id) =>
          client.request(`/api/v2/users/${id}.json`).catch((error) => ({
            user: { id, name: "System Bot", role: "system" },
          }))
        )
      );
      const users = userResponses.map((response) => response.user);
      const userMap = users.reduce((map, user) => {
        map[user.id] = user;
        return map;
      }, {});

      const formattedComments = comments
        .map((comment) => {
          const author = userMap[comment.author_id];
          let authorName = "System Bot";
          let role = "system";

          if (author) {
            authorName = author.name;
            role = author.role === "end-user" ? "Customer" : "Agent";
          }

          const displayName =
            hideCustomerInformation && role === "Customer"
              ? "[Customer]"
              : authorName;

          return `${displayName} (${role}): ${comment.body}`;
        })
        .join("\n");

      const previousMessages = getPreviousMessages();
      const ticketSummary = `
    ### TICKET SUMMARY                
    Zendesk Ticket #${ticketInfo.id}
    Subject: ${ticketInfo.subject}
    Customer Name: ${ticketInfo.customerName}
    Customer Email: ${ticketInfo.customerEmail}
    Status: ${ticketInfo.status}
    Priority: ${ticketInfo.priority}
    Type: ${ticketInfo.type}
    Tags: ${ticketInfo.tags}
    Created At: ${ticketInfo.createdAt}
    Updated At: ${ticketInfo.updatedAt}
    Assignee: ${ticketInfo.assignee} (${ticketInfo.assignee_email})
    Group: ${ticketInfo.group}
    
    Conversation History:
    ${formattedComments}
    ### END TICKET SUMMARY
    
    ### CURRENT CONVERSATION
    ${previousMessages}
    `;

      const payload = {
        message: {
          content: ticketSummary,
          mentions: [
            {
              configurationId: selectedAssistantId,
            },
          ],
          context: {
            username: userFullName.replace(/\s/g, ""),
            timezone: Intl.DateTimeFormat().resolvedOptions().timeZone,
            fullName: userFullName,
            email: userData.currentUser.email,
            profilePictureUrl: "",
            origin: "zendesk",
          },
        },
        title: `Zendesk Ticket #${ticketInfo.id} - ${ticketInfo.customerName}`,
        visibility: "unlisted",
        blocking: false,
        skipToolsValidation: true,
      };

      const options = {
        url: dustApiUrl,
        type: "POST",
        contentType: "application/json",
        headers: {
          Authorization: authorization,
        },
        data: JSON.stringify(payload),
        secure: isProd,
      };

      const response = await client.request(options);

      // Start polling for the conversation result
      await pollConversationEvents(
        response.conversation.sId,
        uniqueId,
        dustWorkspaceId,
        authorization,
        baseUrl,
        metadata
      );

      dustResponse.scrollTop = dustResponse.scrollHeight;

      await client.invoke("resize", { width: "100%", height: "600px" });
    } catch (error) {
      console.error("Error receiving response from Dust:", error);

      const assistantMessageElement = document.getElementById(
        `assistant-${uniqueId}`
      );
      if (assistantMessageElement) {
        assistantMessageElement.innerHTML = `
        <h4>Error:</h4>
        <pre>${
          error.message ||
          "Error receiving response from Dust. Please try again."
        }</pre>
      `;
      }

      dustResponse.scrollTop = dustResponse.scrollHeight;
    } finally {
      userInput.disabled = false;
      sendToDustButton.innerHTML = `
      <svg viewBox="0 0 24 24">
        <path d="M4 12l1.41 1.41L11 7.83V20h2V7.83l5.58 5.59L20 12l-8-8-8 8z"/>
      </svg>
    `;
    }
  }

  async function useAnswer(button, type) {
    try {
      const assistantMessageDiv = button.closest(".assistant-message");
      const answerContent =
        assistantMessageDiv.querySelector(".markdown-content").innerHTML;
      const formattedAnswer = answerContent.replace(/\n/g, "<br>");

      if (type === "private") {
        await client.set("ticket.comment.type", "internalNote");
      } else {
        await client.set("ticket.comment.type", "publicReply");
      }

      await client.invoke("ticket.editor.insert", formattedAnswer);
    } catch (error) {
      console.error(`Error inserting answer as ${type} note:`, error);
    }
  }

  async function copyMessage(button) {
    try {
      const assistantMessageDiv = button.closest(".assistant-message");
      const markdownContent = assistantMessageDiv.querySelector(".markdown-content");

      // Get the original markdown text stored in the data attribute
      const originalMarkdown = markdownContent.dataset.markdown;

      if (originalMarkdown) {
        await navigator.clipboard.writeText(originalMarkdown);
      } else {
        // Fallback: copy the text content
        await navigator.clipboard.writeText(markdownContent.textContent);
      }

      // Visual feedback
      const icon = button.querySelector('svg');
      const originalIcon = icon.innerHTML;

      // Change to checkmark icon
      icon.innerHTML = '<path d="M9 16.17L4.83 12l-1.42 1.41L9 19 21 7l-1.41-1.41z"/>';
      button.classList.add('copied');

      // Reset after 2 seconds
      setTimeout(() => {
        icon.innerHTML = originalIcon;
        button.classList.remove('copied');
      }, 2000);
    } catch (error) {
      console.error('Error copying message:', error);
    }
  }

  function generateUniqueId() {
    return "id-" + Date.now() + "-" + Math.random().toString(36).substr(2, 9);
  }

  function getPreviousMessages() {
    const dustResponse = document.getElementById("dustResponse");
    const messages = dustResponse.getElementsByTagName("div");
    let previousMessages = "";

    for (const messageDiv of messages) {
      const senderElement = messageDiv.querySelector("strong, h4");
      const contentElement = messageDiv.querySelector("pre");

      if (senderElement && contentElement) {
        const sender = senderElement.textContent.trim();
        const content = contentElement.textContent.trim();
        previousMessages += `${sender} ${content}\n\n`;
      }
    }

    return previousMessages;
  }
})();
