import { useEffect, useRef } from "react"
import { Streamlit, withStreamlitConnection, ComponentProps } from "streamlit-component-lib"
import { EditorState } from "@codemirror/state"
import { EditorView, keymap } from "@codemirror/view"
import { defaultKeymap, insertNewlineAndIndent } from "@codemirror/commands"
import { markdown } from "@codemirror/lang-markdown"

function LiveEditor(props: ComponentProps) {
  const containerRef = useRef<HTMLDivElement | null>(null)
  const editorRef = useRef<EditorView | null>(null)
  const suppressNextUpdateRef = useRef(false)

  const incomingText = (props.args.text as string) ?? ""
  const height = (props.args.height as number) ?? 500

  const sendState = (view: EditorView, pressedEnter = false) => {
    const doc = view.state.doc
    const cursor = view.state.selection.main.head
    const line = doc.lineAt(cursor)

    Streamlit.setComponentValue({
      text: doc.toString(),
      cursorLine: line.number - 1,
      cursorCh: cursor - line.from,
      lineCount: doc.lines,
      pressedEnter
    })

    Streamlit.setFrameHeight(height + 10)
  }

  useEffect(() => {
    if (!containerRef.current || editorRef.current) return

    const updateListener = EditorView.updateListener.of((update) => {
      if (!update.docChanged && !update.selectionSet) return

      if (suppressNextUpdateRef.current) {
        suppressNextUpdateRef.current = false
        return
      }

      sendState(update.view, false)
    })

    const enterBinding = keymap.of([
      {
        key: "Enter",
        run(view) {
          insertNewlineAndIndent(view)
          sendState(view, true)
          return true
        }
      }
    ])

    const state = EditorState.create({
      doc: incomingText,
      extensions: [
        markdown(),
        keymap.of(defaultKeymap),
        enterBinding,
        updateListener,
        EditorView.lineWrapping,
        EditorView.theme({
          "&": {
            height: `${height}px`,
            border: "1px solid #ddd",
            borderRadius: "8px",
            fontSize: "16px"
          },
          ".cm-scroller": {
            overflow: "auto"
          },
          ".cm-content": {
            padding: "12px"
          },
          ".cm-focused": {
            outline: "none"
          }
        })
      ]
    })

    editorRef.current = new EditorView({
      state,
      parent: containerRef.current
    })

    sendState(editorRef.current, false)
  }, [height, incomingText])

  useEffect(() => {
    const view = editorRef.current
    if (!view) return

    const currentText = view.state.doc.toString()
    if (currentText === incomingText) return

    suppressNextUpdateRef.current = true
    view.dispatch({
      changes: {
        from: 0,
        to: currentText.length,
        insert: incomingText
      }
    })

    Streamlit.setFrameHeight(height + 10)
  }, [incomingText, height])

  return <div ref={containerRef} />
}

export default withStreamlitConnection(LiveEditor)