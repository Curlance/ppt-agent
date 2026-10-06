window.__ModuleLoader__.load({
  id: '@local/ppt-agent-mcp',
  factory(require) {
    const React = require('react');
    const h = React.createElement;

    const ENDPOINT = 'http://127.0.0.1:8791/strip.txt';
    const OBSERVER = 'http://127.0.0.1:8791/';
    const REFRESH_MS = 3000;

    function StatusStrip() {
      const [text, setText] = React.useState('○ ppt-agent 正在连接…');
      const [down, setDown] = React.useState(false);

      React.useEffect(() => {
        let alive = true;

        async function tick() {
          try {
            const resp = await fetch(ENDPOINT, { cache: 'no-store' });
            if (!resp.ok) throw new Error('HTTP ' + resp.status);
            const line = (await resp.text()).trim();
            if (alive) { setText(line || '○ ppt-agent 无状态'); setDown(false); }
          } catch (err) {
            if (alive) { setText('○ ppt-agent 未连接 —— 先运行 pptctl serve'); setDown(true); }
          }
        }

        tick();
        const timer = setInterval(tick, REFRESH_MS);
        return () => { alive = false; clearInterval(timer); };
      }, []);

      return h(
        'a',
        {
          href: OBSERVER,
          target: '_blank',
          rel: 'noreferrer',
          title: '在浏览器里打开 ppt-agent 观察台',
          style: {
            display: 'block',
            width: '100%',
            boxSizing: 'border-box',
            padding: '2px 4px',
            fontSize: 12,
            lineHeight: '16px',
            textDecoration: 'none',
            // 继承宿主主题：连上了就用前景色，断了才变淡
            color: down ? 'var(--dsw-alias-text-3, #8b93a7)' : 'inherit',
            opacity: down ? 0.75 : 1,
            whiteSpace: 'nowrap',
            overflow: 'hidden',
            textOverflow: 'ellipsis',
            cursor: 'pointer',
          },
        },
        text
      );
    }

    return {
      inject: ['slots'],
      apply(ctx) {
        ctx.slots.inject('conversation.composer.dock', () => ctx.slots.register({
          name: 'conversation.composer.dock',
          id: 'ppt-agent',
          order: 20,
        }, StatusStrip));
      },
    };
  },
});
