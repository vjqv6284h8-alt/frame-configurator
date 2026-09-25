/** Определяет, что текст написан по-русски (кириллица доминирует). */
export function isRussianText(text: string): boolean {
  const trimmed = text.trim()
  if (!trimmed) return false

  const cyrillic = (trimmed.match(/[\u0400-\u04FF]/g) ?? []).length
  const latin = (trimmed.match(/[a-zA-ZÀ-ÿ]/g) ?? []).length
  const letters = cyrillic + latin

  if (letters === 0) return false
  return cyrillic / letters >= 0.5
}

export type TranslationDirection = {
  sourceLang: 'ru' | 'auto'
  targetLang: 'en' | 'ru'
  label: string
}

export function getTranslationDirection(text: string): TranslationDirection {
  if (isRussianText(text)) {
    return {
      sourceLang: 'ru',
      targetLang: 'en',
      label: 'Русский → English',
    }
  }

  return {
    sourceLang: 'auto',
    targetLang: 'ru',
    label: '→ Русский',
  }
}
