export default async function handler(req, res) {
  // Enable CORS for your Excel add-in domain
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET, POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type, Authorization');
  
  // Handle preflight requests
  if (req.method === 'OPTIONS') {
    return res.status(200).end();
  }
  
  // Extract the API path and query parameters
  const { path, ...queryParams } = req.query;
  
  if (!path) {
    return res.status(400).json({ error: 'API path is required' });
  }
  
  // Get authorization header
  const authHeader = req.headers.authorization;
  if (!authHeader) {
    return res.status(401).json({ error: 'Authorization header is required' });
  }
  
  // Determine base URL based on region parameter
  const region = req.headers['x-dust-region'] || queryParams.region;
  const baseUrl = region && region.toLowerCase() === 'eu' 
    ? 'https://eu.dust.tt' 
    : 'https://dust.tt';
  
  // Build the target URL
  const targetUrl = `${baseUrl}${path}`;
  
  try {
    // Prepare fetch options
    const fetchOptions = {
      method: req.method,
      headers: {
        'Authorization': authHeader,
        'Content-Type': 'application/json',
      },
    };
    
    // Add body for POST requests
    if (req.method === 'POST' && req.body) {
      fetchOptions.body = JSON.stringify(req.body);
    }
    
    // Make the request to Dust API
    const response = await fetch(targetUrl, fetchOptions);
    
    // Get the response data
    const data = await response.json();
    
    // Return the response with the same status code
    return res.status(response.status).json(data);
    
  } catch (error) {
    console.error('Proxy error:', error);
    return res.status(500).json({ 
      error: 'Failed to proxy request to Dust API',
      details: error.message 
    });
  }
}