// app/api/forms/route.ts
import { NextResponse } from 'next/server'

export async function GET() {
  try {
    // Mock data for initial testing
    const mockForms = {
      value: [
        {
          id: "form1",
          title: "Test Form 1",
          createdDateTime: "2024-01-15T10:00:00Z",
          responseCount: 5
        },
        {
          id: "form2", 
          title: "Test Form 2",
          createdDateTime: "2024-01-15T11:00:00Z",
          responseCount: 3
        }
      ]
    }

    return NextResponse.json(mockForms)
  } catch (error) {
    console.error('Forms API Error:', error)
    return NextResponse.json(
      { error: 'Internal Server Error' },
      { status: 500 }
    )
  }
}

// Optional: Add POST handler if needed
export async function POST(request: Request) {
  try {
    const body = await request.json()
    
    return NextResponse.json({
      message: 'Form data received',
      data: body
    })
  } catch (error) {
    console.error('Forms API Error:', error)
    return NextResponse.json(
      { error: 'Internal Server Error' },
      { status: 500 }
    )
  }
}

// Types for better type safety
export interface Form {
  id: string
  title: string
  createdDateTime: string
  responseCount: number
}

export interface FormsResponse {
  value: Form[]
}