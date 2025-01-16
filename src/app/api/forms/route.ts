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

// Logger utility
const logger = {
  info: (message: string, data?: any) => {
    console.log(`[INFO] ${message}`, data ? JSON.stringify(data, null, 2) : '')
  },
  error: (message: string, error: any) => {
    console.error(`[ERROR] ${message}`, error)
    if (error.response) {
      console.error('[ERROR] Response:', {
        status: error.response.status,
        headers: error.response.headers,
        data: error.response.data
      })
    }
  },
  debug: (message: string, data?: any) => {
    console.log(`[DEBUG] ${message}`, data ? JSON.stringify(data, null, 2) : '')
  }
}

// Graph API client configuration
const msalConfig = {
  auth: {
    clientId: process.env.AZURE_CLIENT_ID!,
    clientSecret: process.env.AZURE_CLIENT_SECRET!,
    authority: `https://login.microsoftonline.com/${process.env.AZURE_TENANT_ID}`
  }
}

const FORM_ID = process.env.NEXT_PUBLIC_FORM_ID!;

// Initialize MSAL application
const cca = new ConfidentialClientApplication(msalConfig)

async function getAccessToken() {
  logger.debug('Getting access token...')
  try {
    const tokenRequest = {
      scopes: ['https://graph.microsoft.com/.default']
    }
    logger.debug('Token request', tokenRequest)
    
    const response = await cca.acquireTokenByClientCredential(tokenRequest)
    logger.info('Successfully acquired access token', {
      tokenLength: response?.accessToken?.length,
      expiresOn: response?.expiresOn
    })
    return response?.accessToken
  } catch (error) {
    logger.error('Failed to acquire token', error)
    throw error
  }
}

async function getAllOrgForms() {
  logger.debug('Fetching all organization forms...')
  const accessToken = await getAccessToken()
  
  try {
    // Using the organization forms endpoint
    const response = await fetch(
      'https://graph.microsoft.com/v1.0/organization/settings/form/forms',
      {
        headers: {
          'Authorization': `Bearer ${accessToken}`,
          'Content-Type': 'application/json'
        }
      }
    )

    logger.debug('Forms API response status', {
      status: response.status,
      statusText: response.statusText,
      headers: Object.fromEntries(response.headers.entries())
    })

    if (!response.ok) {
      const errorText = await response.text()
      logger.error('Forms API error response', {
        status: response.status,
        statusText: response.statusText,
        error: errorText
      })
      throw new Error(`Failed to fetch forms: ${response.statusText}`)
    }

    const data = await response.json()
    logger.info('Successfully fetched organization forms', {
      count: data.value?.length,
      forms: data.value?.map((f: any) => ({
        id: f.id,
        title: f.title
      }))
    })
    return data
  } catch (error) {
    logger.error('Error in getAllOrgForms', error)
    throw error
  }
}

async function fetchFormResponses(formId: string) {
  logger.debug('Fetching form responses...', { formId })
  const accessToken = await getAccessToken()
  
  try {
    // Using the organization forms endpoint for responses
    const url = `https://graph.microsoft.com/v1.0/organization/settings/form/forms/${formId}/responses`
    logger.debug('Making request to', { url })

    const response = await fetch(
      url,
      {
        headers: {
          'Authorization': `Bearer ${accessToken}`,
          'Content-Type': 'application/json'
        }
      }
    )

    logger.debug('Form responses API response status', {
      status: response.status,
      statusText: response.statusText,
      headers: Object.fromEntries(response.headers.entries())
    })

    if (!response.ok) {
      const errorText = await response.text()
      logger.error('Form responses API error', {
        status: response.status,
        statusText: response.statusText,
        error: errorText
      })
      throw new Error(`Failed to fetch form responses: ${response.statusText}`)
    }

    const data = await response.json()
    logger.info('Successfully fetched form responses', {
      responseCount: data.value?.length
    })
    return data
  } catch (error) {
    logger.error('Error in fetchFormResponses', error)
    throw error
  }
}

export async function GET() {
  logger.info('Starting GET request handler')
  
  try {
    logger.debug('Using form ID', { FORM_ID })

    // First, let's get all organization forms
    logger.info('Fetching all organization forms...')
    const allForms = await getAllOrgForms()
    
    // Find the form that matches our ID
    logger.debug('Searching for matching form', { 
      searchId: FORM_ID,
      availableForms: allForms.value?.map((f: any) => ({ 
        id: f.id, 
        title: f.title 
      }))
    })

    const form = allForms.value?.find((f: any) => {
      const matches = f.id.includes(FORM_ID) || FORM_ID.includes(f.id)
      logger.debug('Checking form match', { 
        formId: f.id, 
        matches,
        title: f.title
      })
      return matches
    })

    if (!form) {
      logger.error('Form not found', { searchId: FORM_ID })
      return NextResponse.json(
        { error: 'Form not found' },
        { status: 404 }
      )
    }

    logger.info('Found matching form', { 
      id: form.id, 
      title: form.title 
    })

    const formResponses = await fetchFormResponses(form.id)
    
    // Transform responses
    logger.debug('Transforming responses', { 
      responseCount: formResponses.value?.length 
    })

    const transformedResponses = formResponses.value?.map((response: FormResponse) => {
      const transformed = {
        id: response.id,
        submitDate: response.submitDate,
        respondent: response.respondent.emailAddress,
        answers: response.answers
      }
      logger.debug('Transformed response', { 
        id: transformed.id,
        submitDate: transformed.submitDate,
        answerCount: transformed.answers.length
      })
      return transformed
    })

    logger.info('Successfully processed all responses', { 
      count: transformedResponses?.length 
    })

    return NextResponse.json({ responses: transformedResponses })
  } catch (error) {
    logger.error('Error in GET handler', error)
    return NextResponse.json(
      { 
        error: 'Failed to fetch form responses', 
        details: error instanceof Error ? error.message : 'Unknown error'
      },
      { status: 500 }
    )
  }
}