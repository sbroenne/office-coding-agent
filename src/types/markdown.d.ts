/** Allow importing .md files as raw strings (via Vite md-raw plugin). */
declare module '*.md' {
  const content: string;
  export default content;
}

/** Allow importing .md files with Vite's ?raw suffix. */
declare module '*.md?raw' {
  const content: string;
  export default content;
}
