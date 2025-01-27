// app/api/excel/route.ts
import { NextRequest, NextResponse } from 'next/server';
import axios from 'axios';

const CLIENT_ID = process.env.AZURE_CLIENT_ID;
const CLIENT_SECRET = process.env.AZURE_CLIENT_SECRET;
const TENANT_ID = process.env.AZURE_TENANT_ID;
const FILE_ID = '6f966bfb-10c3-47cd-932c-d51fefdb5ba2';
// You'll need to provide your organization's site ID
const SITE_ID = 'root'; // or your specific SharePoint site ID
// You'll need your organization's drive ID
const DRIVE_ID = 'b!6f966bfb-10c3-47cd-932c-d51fefdb5ba2'; // or your specific drive ID

async function getAccessToken(): Promise<string> {
  try {
    console.log('Attempting to get access token...');
    
    const tokenEndpoint = `https://login.microsoftonline.com/${TENANT_ID}/oauth2/v2.0/token`;
    const params = new URLSearchParams({
      client_id: CLIENT_ID!,
      client_secret: CLIENT_SECRET!,
      grant_type: 'client_credentials',
      scope: 'https://graph.microsoft.com/.default'
    });

    const response = await axios.post(tokenEndpoint, params.toString(), {
      headers: {
        'Content-Type': 'application/x-www-form-urlencoded'
      }
    });

    console.log('Token acquired successfully');
    return response.data.access_token;

  } catch (error) {
    if (axios.isAxiosError(error)) {
      console.error('Token Error:', error.response?.data);
      throw new Error(`Authentication failed: ${error.response?.data?.error_description || error.message}`);
    }
    throw error;
  }
}

async function getExcelData(accessToken: string) {
  try {
    console.log('Setting up Graph API client...');
    
    const graphApi = axios.create({
      baseURL: 'https://graph.microsoft.com/v1.0',
      headers: {
        'Authorization': `Bearer ${accessToken}`,
        'Content-Type': 'application/json'
      }
    });

    // First, get the drive item to ensure we have access
    console.log('Fetching drive item...');
    const driveItemPath = `/sites/${SITE_ID}/drives/${DRIVE_ID}/items/${FILE_ID}`;
    await graphApi.get(driveItemPath);

    console.log('Fetching worksheets...');
    const worksheetsResponse = await graphApi.get(
      `/sites/${SITE_ID}/drives/${DRIVE_ID}/items/${FILE_ID}/workbook/worksheets`
    );
    
    const firstWorksheetId = worksheetsResponse.data.value[0].id;
    console.log('First worksheet ID:', firstWorksheetId);

    console.log('Fetching worksheet data...');
    const rangeResponse = await graphApi.get(
      `/sites/${SITE_ID}/drives/${DRIVE_ID}/items/${FILE_ID}/workbook/worksheets/${firstWorksheetId}/usedRange`
    );

    return rangeResponse.data.values;
  } catch (error) {
    if (axios.isAxiosError(error)) {
      console.error('Graph API error:', {
        status: error.response?.status,
        statusText: error.response?.statusText,
        data: error.response?.data
      });
      throw new Error(`Failed to fetch Excel data: ${error.response?.data?.error?.message || error.message}`);
    }
    throw error;
  }
}

export async function GET(request: NextRequest) {
  try {
    console.log('Starting Excel data fetch process...');
    
    const accessToken = await getAccessToken();
    console.log('Successfully obtained access token');
    
    const data = await getExcelData(accessToken);
    console.log('Successfully fetched Excel data');

    if (!data || data.length === 0) {
      return NextResponse.json({
        success: false,
        error: 'No data found in Excel file'
      }, { status: 404 });
    }

    // Convert array of arrays to array of objects
    const [headers, ...rows] = data;
    const formattedData = rows.map((row: any[]) => 
      Object.fromEntries(headers.map((header: string, index: number) => [header, row[index]]))
    );

    return NextResponse.json({
      success: true,
      data: formattedData
    });

  } catch (error: any) {
    console.error('API Error:', {
      message: error.message,
      response: error.response?.data
    });
    
    return NextResponse.json({
      success: false,
      error: error.message || 'An unexpected error occurred',
      details: process.env.NODE_ENV === 'development' ? {
        message: error.message,
        response: error.response?.data
      } : undefined
    }, { status: 500 });
  }
}