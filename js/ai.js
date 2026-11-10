// Claude API 연동 (선택 기능): 설정에 API 키를 넣으면 AI 요약·질문이 활성화됩니다.
// 키는 이 브라우저의 localStorage에만 저장되며, 요청은 브라우저에서 api.anthropic.com으로 바로 전송됩니다.
import { settings } from './store.js';

const SDK_URL = './vendor/anthropic-sdk.mjs';
let clientCache = null;

async function getClient() {
  const { apiKey } = settings.get();
  if (!apiKey) throw new Error('설정 탭에서 Claude API 키를 먼저 입력해주세요.');
  if (clientCache?.key === apiKey) return clientCache.client;
  const { default: Anthropic } = await import(SDK_URL);
  // 개인용 앱이라 브라우저에서 직접 호출합니다(키가 이 기기에만 저장됨).
  const client = new Anthropic({ apiKey, dangerouslyAllowBrowser: true });
  clientCache = { key: apiKey, client };
  return client;
}

export const aiEnabled = () => !!settings.get().apiKey;

async function ask({ system, prompt, schema, maxTokens = 16000 }) {
  const client = await getClient();
  const { model } = settings.get();
  const params = {
    model: model || 'claude-opus-5-5',
    max_tokens: maxTokens,
    system,
    output_config: { effort: 'medium', ...(schema ? { format: { type: 'json_schema', schema } } : {}) },
    // 안전 정책으로 거절될 경우 서버가 자동으로 대체 모델로 재시도
    betas: ['server-side-fallback-2026-07-01'],
    fallbacks: 'default',
    messages: [{ role: 'user', content: prompt }],
  };
  let res;
  try {
    res = await client.beta.messages.create(params);
  } catch (e) {
    const status = e?.status;
    if (status === 401) throw new Error('API 키가 올바르지 않습니다. 설정을 확인해주세요.');
    if (status === 429) throw new Error('요청이 너무 많습니다. 잠시 후 다시 시도해주세요.');
    if (status >= 500) throw new Error('Claude 서버가 일시적으로 응답하지 않습니다. 잠시 후 다시 시도해주세요.');
    throw new Error('AI 요청 실패: ' + (e?.message || e));
  }
  if (res.stop_reason === 'refusal') throw new Error('AI가 이 요청을 처리하지 않았습니다.');
  const text = res.content.filter((b) => b.type === 'text').map((b) => b.text).join('');
  return schema ? JSON.parse(text) : text;
}

const SUMMARY_SCHEMA = {
  type: 'object',
  additionalProperties: false,
  required: ['summary', 'decisions', 'actions', 'keywords'],
  properties: {
    summary: { type: 'string' },
    decisions: { type: 'array', items: { type: 'string' } },
    actions: {
      type: 'array',
      items: {
        type: 'object',
        additionalProperties: false,
        required: ['text', 'owner', 'due'],
        properties: {
          text: { type: 'string' },
          owner: { type: 'string' },
          due: { type: 'string' },
        },
      },
    },
    keywords: { type: 'array', items: { type: 'string' } },
  },
};

export async function aiSummarize(meeting) {
  return ask({
    system:
      '당신은 회사 회의록을 정리하는 비서입니다. 회의 메모를 읽고 한국어로 정리합니다. ' +
      'summary는 3~5문장의 핵심 요약, decisions는 확정된 결정사항, actions는 후속 할일(담당자 owner, 마감일 due는 YYYY-MM-DD, 모르면 빈 문자열), ' +
      'keywords는 나중에 검색에 쓸 핵심 키워드 3~6개입니다. 메모에 없는 내용은 지어내지 마세요.',
    prompt:
      `회의 날짜: ${meeting.date}\n회의 제목: ${meeting.title}\n참석자: ${meeting.attendees || '-'}\n\n` +
      `<memo>\n${meeting.content}\n</memo>`,
    schema: SUMMARY_SCHEMA,
  });
}

export async function aiAnswer(question, contextText) {
  return ask({
    system:
      '당신은 사용자의 업무 기록(일정, 회의 메모, 할일)을 바탕으로 질문에 답하는 업무 비서입니다. ' +
      '제공된 기록에 근거해서만 한국어로 간결하게 답하고, 관련 날짜와 회의 제목을 함께 알려주세요. ' +
      '표로 정리하면 좋은 내용은 마크다운 표로 작성하세요. 기록에 없으면 없다고 말하세요.',
    prompt: `<records>\n${contextText}\n</records>\n\n질문: ${question}`,
  });
}
