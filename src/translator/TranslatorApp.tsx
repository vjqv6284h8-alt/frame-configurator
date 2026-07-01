import { useEffect, useRef, useState } from 'react'
import type { FormEvent, KeyboardEvent } from 'react'
import { getTranslationDirection } from './detectLanguage'
import './TranslatorApp.css'

type ChatMessage = {
  id: string
  original: string
  translation: string
  direction: string
  createdAt: string
  error?: string
}

type TranslateResponse = {
  original: string
  translation: string
  direction: string
  sourceLang: string
  targetLang: string
}

type TranslateErrorResponse = {
  message?: string
  error?: string
}

const createId = () =>
  typeof crypto !== 'undefined' && 'randomUUID' in crypto
    ? crypto.randomUUID()
    : `${Date.now()}-${Math.random().toString(16).slice(2)}`

export default function TranslatorApp() {
  const [messages, setMessages] = useState<ChatMessage[]>([])
  const [input, setInput] = useState('')
  const [isLoading, setIsLoading] = useState(false)
  const chatRef = useRef<HTMLDivElement>(null)
  const inputRef = useRef<HTMLTextAreaElement>(null)

  const previewDirection = input.trim() ? getTranslationDirection(input).label : 'Русский ↔ другие языки'

  useEffect(() => {
    chatRef.current?.scrollTo({ top: chatRef.current.scrollHeight, behavior: 'smooth' })
  }, [messages, isLoading])

  const sendMessage = async () => {
    const text = input.trim()
    if (!text || isLoading) return

    const messageId = createId()
    const createdAt = new Date().toLocaleTimeString('ru-RU', {
      hour: '2-digit',
      minute: '2-digit',
    })

    setInput('')
    setIsLoading(true)
    setMessages((prev) => [
      ...prev,
      {
        id: messageId,
        original: text,
        translation: '…',
        direction: getTranslationDirection(text).label,
        createdAt,
      },
    ])

    try {
      const response = await fetch('/api/translate', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ text }),
      })

      const payload = (await response.json()) as TranslateResponse & TranslateErrorResponse

      if (!response.ok) {
        throw new Error(payload.message || payload.error || 'Ошибка перевода')
      }

      setMessages((prev) =>
        prev.map((message) =>
          message.id === messageId
            ? {
                ...message,
                translation: payload.translation,
                direction: payload.direction,
              }
            : message,
        ),
      )
    } catch (error) {
      const errorText = error instanceof Error ? error.message : 'Неизвестная ошибка'
      setMessages((prev) =>
        prev.map((message) =>
          message.id === messageId
            ? {
                ...message,
                translation: '',
                error: errorText,
              }
            : message,
        ),
      )
    } finally {
      setIsLoading(false)
      inputRef.current?.focus()
    }
  }

  const handleSubmit = (event: FormEvent) => {
    event.preventDefault()
    void sendMessage()
  }

  const handleKeyDown = (event: KeyboardEvent<HTMLTextAreaElement>) => {
    if (event.key === 'Enter' && !event.shiftKey) {
      event.preventDefault()
      void sendMessage()
    }
  }

  return (
    <div className="translator-app">
      <header className="translator-header">
        <h1>Чат-переводчик</h1>
        <p>Русский → English · любой другой язык → Русский</p>
      </header>

      <div className="translator-chat" ref={chatRef}>
        {messages.length === 0 ? (
          <div className="translator-empty">
            Напишите сообщение — перевод появится здесь. Enter отправляет, Shift+Enter — новая строка.
          </div>
        ) : (
          messages.map((message) => (
            <article className="translator-message" key={message.id}>
              <div className="translator-message-meta">
                <span className="translator-message-direction">{message.direction}</span>
                <time>{message.createdAt}</time>
              </div>
              <p className="translator-message-original">{message.original}</p>
              {message.error ? (
                <p className="translator-message-error">{message.error}</p>
              ) : (
                <p className="translator-message-translation">{message.translation}</p>
              )}
            </article>
          ))
        )}
      </div>

      <form className="translator-composer" onSubmit={handleSubmit}>
        <div className="translator-composer-inner">
          <textarea
            ref={inputRef}
            className="translator-input"
            value={input}
            onChange={(event) => setInput(event.target.value)}
            onKeyDown={handleKeyDown}
            placeholder="Введите текст для перевода…"
            rows={2}
            disabled={isLoading}
          />
          <button className="translator-send" type="submit" disabled={isLoading || !input.trim()}>
            {isLoading ? '…' : 'Перевести'}
          </button>
        </div>
        <p className="translator-hint">Направление: {previewDirection}</p>
      </form>
    </div>
  )
}
