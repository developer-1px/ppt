import { MessageSquare } from 'lucide-react'
import { useState } from 'react'
import { PPT_COMMENT_THREAD_MODEL } from '../../pptCanvasAppAffordanceAdapter'
import {
  getPPTCommentThread,
  normalizePPTCommentReplyBody,
  PPT_COMMENT_BODY_MAX_LENGTH,
  PPT_COMMENT_REPLY_MAX_LENGTH,
} from '../../pptCommentThreadAdapter'
import { Button } from '../core'
import { createPPTInspectorActionDispatcher } from './PPTInspectorActionDispatcher'
import type { PPTInspectorProps } from './PPTInspectorContract'

type PPTCommentInspectorFieldsProps = Pick<
  PPTInspectorProps,
  'model' | 'onAction'
>

export function PPTCommentInspectorFields({
  model,
  onAction,
}: PPTCommentInspectorFieldsProps) {
  const comment = model.selectedElement
  const {
    onCommentBodyChange,
    onCommentReplyAdd,
    onCommentResolvedChange,
  } = createPPTInspectorActionDispatcher(onAction)
  const [replyDraftById, setReplyDraftById] =
    useState<Record<string, string>>({})

  if (comment?.kind !== 'comment') {
    return null
  }

  const commentId = comment.id
  const thread = getPPTCommentThread(comment)
  const replyDraft = replyDraftById[commentId] ?? ''

  function updateReplyDraft(value: string) {
    setReplyDraftById((current) => ({
      ...current,
      [commentId]: normalizePPTCommentReplyBody(value),
    }))
  }

  function commitReplyDraft() {
    if (replyDraft.trim().length === 0) {
      return
    }

    onCommentReplyAdd(commentId, replyDraft)
    setReplyDraftById((current) => ({
      ...current,
      [commentId]: '',
    }))
  }

  return (
    <>
      <label className="ppt-field">
        <span>Comment</span>
        <textarea
          data-ppt-style-field="comment-body"
          maxLength={PPT_COMMENT_BODY_MAX_LENGTH}
          value={comment.body}
          onChange={(event) =>
            onCommentBodyChange(comment.id, event.target.value)}
        />
      </label>
      <label className="ppt-checkbox-field">
        <input
          checked={comment.resolved === true}
          data-ppt-style-field="comment-resolved"
          type="checkbox"
          onChange={(event) =>
            onCommentResolvedChange(comment.id, event.target.checked)}
        />
        <span>Resolved</span>
      </label>
      <section
        className="ppt-comment-thread"
        data-ppt-comment-thread
        data-ppt-comment-thread-count={thread.length}
        data-ppt-comment-thread-model={PPT_COMMENT_THREAD_MODEL}
        data-ppt-comment-thread-resolved={
          comment.resolved === true ? 'true' : 'false'
        }
      >
        <div className="ppt-comment-thread-header">
          <span>Thread</span>
          <span data-ppt-comment-thread-count-label>{thread.length}</span>
        </div>
        <div className="ppt-comment-thread-list">
          {thread.map((message, index) => (
            <article
              className="ppt-comment-thread-message"
              data-ppt-comment-thread-message={message.id}
              data-ppt-comment-thread-message-index={index}
              key={message.id}
            >
              <div className="ppt-comment-thread-meta">
                <span data-ppt-comment-thread-author>{message.authorName}</span>
                <span data-ppt-comment-thread-created>{message.createdAt}</span>
              </div>
              <p data-ppt-comment-thread-body>{message.body}</p>
            </article>
          ))}
        </div>
        <div className="ppt-comment-reply-row">
          <label className="ppt-field">
            <span>Reply</span>
            <textarea
              data-ppt-comment-reply-input
              data-ppt-style-field="comment-reply"
              maxLength={PPT_COMMENT_REPLY_MAX_LENGTH}
              value={replyDraft}
              onChange={(event) => updateReplyDraft(event.target.value)}
            />
          </label>
          <Button
            data-ppt-comment-reply-add
            disabled={replyDraft.trim().length === 0}
            onClick={commitReplyDraft}
          >
            <MessageSquare size={15} /> Add
          </Button>
        </div>
      </section>
    </>
  )
}
