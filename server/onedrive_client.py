import os
import aiohttp
import asyncio
from typing import List, Dict, Any, Optional
import json

class OneDriveClient:
    """Client for interacting with OneDrive via Microsoft Graph API"""
    
    def __init__(self):
        self.graph_url = "https://graph.microsoft.com/v1.0"
        self.connection_settings = None
    
    async def _get_access_token(self) -> str:
        """Get a valid access token for Microsoft Graph API"""
        # Check if we have a cached token that's still valid
        if (self.connection_settings and 
            self.connection_settings.get('settings', {}).get('expires_at') and
            self._is_token_valid(self.connection_settings['settings']['expires_at'])):
            return self.connection_settings['settings']['access_token']
        
        # Get fresh token from Replit connector
        hostname = os.getenv('REPLIT_CONNECTORS_HOSTNAME')
        x_replit_token = None
        
        repl_identity = os.getenv('REPL_IDENTITY')
        web_repl_renewal = os.getenv('WEB_REPL_RENEWAL')
        
        if repl_identity:
            x_replit_token = 'repl ' + repl_identity
        elif web_repl_renewal:
            x_replit_token = 'depl ' + web_repl_renewal
        
        if not x_replit_token:
            raise Exception('X_REPLIT_TOKEN not found for repl/depl')
        
        url = f'https://{hostname}/api/v2/connection?include_secrets=true&connector_names=onedrive'
        headers = {
            'Accept': 'application/json',
            'X_REPLIT_TOKEN': x_replit_token
        }
        
        async with aiohttp.ClientSession() as session:
            async with session.get(url, headers=headers) as response:
                if response.status != 200:
                    raise Exception(f'Failed to get connection: {response.status}')
                
                data = await response.json()
                self.connection_settings = data.get('items', [{}])[0]
        
        access_token = (self.connection_settings.get('settings', {}).get('access_token') or 
                       self.connection_settings.get('settings', {}).get('oauth', {}).get('credentials', {}).get('access_token'))
        
        if not self.connection_settings or not access_token:
            raise Exception('OneDrive not connected')
        
        return access_token
    
    def _is_token_valid(self, expires_at: str) -> bool:
        """Check if token is still valid"""
        try:
            from datetime import datetime
            expiry = datetime.fromisoformat(expires_at.replace('Z', '+00:00'))
            return expiry.timestamp() > datetime.now().timestamp()
        except:
            return False
    
    async def _make_graph_request(self, endpoint: str, method: str = 'GET', data: Optional[bytes] = None) -> Dict[str, Any]:
        """Make authenticated request to Microsoft Graph API"""
        access_token = await self._get_access_token()
        
        headers = {
            'Authorization': f'Bearer {access_token}',
            'Content-Type': 'application/json' if method != 'PUT' else 'application/octet-stream'
        }
        
        url = f"{self.graph_url}{endpoint}"
        
        async with aiohttp.ClientSession() as session:
            async with session.request(method, url, headers=headers, data=data) as response:
                if response.status >= 400:
                    error_text = await response.text()
                    raise Exception(f'Graph API error {response.status}: {error_text}')
                
                if method == 'GET':
                    return await response.json()
                else:
                    return {}
    
    async def list_word_documents(self) -> List[Dict[str, Any]]:
        """List all Word documents in OneDrive"""
        try:
            # Search for .docx files
            endpoint = "/me/drive/root/search(q='.docx')"
            response = await self._make_graph_request(endpoint)
            
            word_docs = []
            for item in response.get('value', []):
                if (item.get('file') and 
                    item.get('name', '').lower().endswith('.docx') and
                    not item.get('name', '').startswith('~')):  # Skip temp files
                    
                    word_docs.append({
                        'id': item['id'],
                        'name': item['name'],
                        'size': item.get('size', 0),
                        'lastModified': item.get('lastModifiedDateTime'),
                        'webUrl': item.get('webUrl'),
                        'downloadUrl': item.get('@microsoft.graph.downloadUrl')
                    })
            
            return sorted(word_docs, key=lambda x: x.get('lastModified', ''), reverse=True)
            
        except Exception as e:
            raise Exception(f"Failed to list Word documents: {str(e)}")
    
    async def download_document(self, file_id: str) -> bytes:
        """Download a Word document by file ID"""
        try:
            endpoint = f"/me/drive/items/{file_id}/content"
            access_token = await self._get_access_token()
            
            headers = {'Authorization': f'Bearer {access_token}'}
            url = f"{self.graph_url}{endpoint}"
            
            async with aiohttp.ClientSession() as session:
                async with session.get(url, headers=headers) as response:
                    if response.status >= 400:
                        error_text = await response.text()
                        raise Exception(f'Download failed {response.status}: {error_text}')
                    
                    return await response.read()
                    
        except Exception as e:
            raise Exception(f"Failed to download document: {str(e)}")
    
    async def upload_document(self, file_id: str, content: bytes) -> Dict[str, Any]:
        """Upload/update a Word document"""
        try:
            endpoint = f"/me/drive/items/{file_id}/content"
            return await self._make_graph_request(endpoint, method='PUT', data=content)
            
        except Exception as e:
            raise Exception(f"Failed to upload document: {str(e)}")
    
    async def get_file_info(self, file_id: str) -> Dict[str, Any]:
        """Get detailed information about a file"""
        try:
            endpoint = f"/me/drive/items/{file_id}"
            return await self._make_graph_request(endpoint)
            
        except Exception as e:
            raise Exception(f"Failed to get file info: {str(e)}")
