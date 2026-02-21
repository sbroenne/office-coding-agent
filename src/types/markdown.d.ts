/** Allow importing .md files as raw strings (via Vite md-raw plugin). */
declare module '*.md' {
  const content: string;
  export default content;
}
