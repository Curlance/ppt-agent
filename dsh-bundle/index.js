/** ppt-agent 插件的宿主半边。
 *
 * 这里刻意什么都不做：
 * - 工具由 @deepseek-ai/dsh-mcp-client 提供（守护进程独占 COM，见 README）
 * - 界面全部在浏览器半边（client.js），由 package.json 的 dsh.client 装载
 *
 * 空的 apply 意味着这个插件在宿主启动阶段不会引入任何副作用或失败点。
 */
export function apply() {}
