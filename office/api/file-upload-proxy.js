export default async function handler(req, res) {
  // Enable CORS
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type, Authorization');

  // Handle preflight requests
  if (req.method === 'OPTIONS') {
    return res.status(200).end();
  }

  if (req.method !== 'POST') {
    return res.status(405).json({ error: 'Method not allowed' });
  }

  try {
    // Get the upload URL and file data from the request body
    const { uploadUrl, fileContent, fileName, contentType } = req.body;

    if (!uploadUrl || !fileContent || !fileName || !contentType) {
      return res.status(400).json({
        error: 'Missing required fields: uploadUrl, fileContent, fileName, contentType'
      });
    }

    // Get the authorization token
    const authHeader = req.headers.authorization;
    if (!authHeader || !authHeader.startsWith('Bearer ')) {
      return res.status(401).json({ error: 'Missing or invalid authorization header' });
    }

    // Use native FormData (Node 18+ built-in)
    const { FormData, File } = await import('node:buffer');
    const formData = new FormData();

    // Convert text content to buffer
    let buffer;
    if (contentType === 'text/plain') {
      buffer = Buffer.from(fileContent, 'utf-8');
    } else {
      // Assume base64 for other content types
      buffer = Buffer.from(fileContent, 'base64');
    }

    // Create a File object and append to FormData
    const file = new File([buffer], fileName, { type: contentType });
    formData.append('file', file);

    // Upload the file to Dust's upload URL
    const uploadResponse = await fetch(uploadUrl, {
      method: 'POST',
      headers: {
        'Authorization': authHeader,
      },
      body: formData,
    });

    // Get the response
    const responseData = await uploadResponse.json().catch(() => ({}));

    if (!uploadResponse.ok) {
      return res.status(uploadResponse.status).json({
        error: responseData?.error?.message || `Upload failed with status ${uploadResponse.status}`,
        details: responseData
      });
    }

    // Return the successful upload response
    return res.status(200).json(responseData);

  } catch (error) {
    console.error('File upload proxy error:', error);
    return res.status(500).json({
      error: 'Failed to proxy file upload',
      details: error.message
    });
  }
}
