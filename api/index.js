export default async function handler(req) {
  const userAgent = req.headers.get('user-agent') || '';
  const githubRepo = 'https://github.com/Louchatfroff/Office-Unnatended-Install';
  const scriptUrl = 'https://raw.githubusercontent.com/Louchatfroff/Office-Unnatended-Install/main/start.ps1';

  // Check if request is from PowerShell, curl, wget, or similar CLI tools
  const isCLI = /powershell|curl|wget|httpie|invoke-webrequest|invoke-restmethod/i.test(userAgent);

  if (isCLI) {
    // Fetch and return the PowerShell script
    try {
      const response = await fetch(scriptUrl);
      const script = await response.text();

      return new Response(script, {
        status: 200,
        headers: {
          'Content-Type': 'text/plain; charset=utf-8',
          'Cache-Control': 'no-cache, no-store, must-revalidate',
        },
      });
    } catch (error) {
      return new Response('Error fetching script', { status: 500 });
    }
  } else {
    // Redirect browsers to GitHub
    return Response.redirect(githubRepo, 302);
  }
}

export const config = {
  runtime: 'edge',
};
