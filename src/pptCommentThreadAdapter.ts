import type { PPTComment, PPTCommentThreadMessage } from './pptModel'

export const PPT_COMMENT_DEFAULT_BODY = 'Comment'
export const PPT_COMMENT_DEFAULT_AUTHOR = 'You'
export const PPT_COMMENT_DEFAULT_CREATED_AT = 'Just now'
export const PPT_COMMENT_BODY_MAX_LENGTH = 240
export const PPT_COMMENT_REPLY_MAX_LENGTH = 240

export function normalizePPTCommentBody(value: string) {
  return value.slice(0, PPT_COMMENT_BODY_MAX_LENGTH)
}

export function normalizePPTCommentReplyBody(value: string) {
  return value.slice(0, PPT_COMMENT_REPLY_MAX_LENGTH)
}

export function normalizePPTCommentAuthorName(value: string) {
  return value.trim().slice(0, 80) || PPT_COMMENT_DEFAULT_AUTHOR
}

export function normalizePPTCommentCreatedAt(value: string) {
  return value.trim().slice(0, 80) || PPT_COMMENT_DEFAULT_CREATED_AT
}

export function normalizePPTCommentMessageId(value: string) {
  return value.trim().slice(0, 120) || 'json-comment:message'
}

export function getPPTCommentThread(
  comment: PPTComment,
): PPTCommentThreadMessage[] {
  if (comment.thread && comment.thread.length > 0) {
    return comment.thread
  }

  return [{
    authorName: comment.authorName ?? PPT_COMMENT_DEFAULT_AUTHOR,
    body: comment.body,
    createdAt: comment.createdAt ?? PPT_COMMENT_DEFAULT_CREATED_AT,
    id: `${comment.id}:message-1`,
  }]
}

export function getPPTCommentThreadWithBody(
  comment: PPTComment,
  body: string,
): PPTComment['thread'] {
  const [first, ...rest] = getPPTCommentThread(comment)

  return [{
    ...first,
    body,
  }, ...rest]
}
