// app/api/forms/route.ts
import { NextResponse } from 'next/server'
import { ConfidentialClientApplication } from '@azure/msal-node'

// Types
interface FormResponse {
  id: string
  respondent: {
    emailAddress: string
  }
  submitDate: string
  answers: Array<{
    questionId: string
    value: string | string[]
  }>
}

// Graph API client configuration
const msalConfig = {
  auth: {
    clientId: process.env.AZURE_CLIENT_ID!,
    clientSecret: process.env.AZURE_CLIENT_SECRET!,
    authority: `https://login.microsoftonline.com/${process.env.AZURE_TENANT_ID}`
  }
}

const FORMID = process.env.NEXT_PUBLIC_FORM_ID!;

// Initialize MSAL application
const cca = new ConfidentialClientApplication(msalConfig)

async function getAccessToken() {
  try {
    const tokenRequest = {
      scopes: ['https://graph.microsoft.com/.default']
    }
    const response = await cca.acquireTokenByClientCredential(tokenRequest)
    return response?.accessToken
  } catch (error) {
    console.error('Error acquiring token:', error)
    throw error
  }
}

async function fetchFormResponses(formId: string) {
  const accessToken = await getAccessToken()
  
  const response = await fetch(
    `https://graph.microsoft.com/v1.0/organization/settings/form/forms/${formId}/responses`,
    {
      headers: {
        'Authorization': `Bearer ${accessToken}`,
        'Content-Type': 'application/json'
      }
    }
  )
  console.log("RES: ", response);

  if (!response.ok) {
    const errorText = await response.text()
    console.error('Error response:', errorText)
    throw new Error(`Failed to fetch form responses: ${response.statusText}`)
  }

  return response.json()
}

export async function GET() {
  try {
    const formId = FORMID;
    console.log("FORM ID: ", formId);

    if (!formId) {
      return NextResponse.json(
        { error: 'Form ID is required' },
        { status: 400 }
      )
    }

    const formResponses = await fetchFormResponses(formId)
    console.log("Form Responses:", formResponses);

    const transformedResponses = formResponses.value.map((response: FormResponse) => ({
      id: response.id,
      submitDate: response.submitDate,
      respondent: response.respondent.emailAddress,
      answers: response.answers
    }))

    return NextResponse.json({ responses: transformedResponses })
  } catch (error) {
    console.error('Forms API Error:', error)
    return NextResponse.json(
      { error: 'Failed to fetch form responses' },
      { status: 500 }
    )
  }
}