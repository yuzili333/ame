// Ignore dates in content fields, but resume parsing at an explicit scheduling
// field. A title before the time must not discard the rest of the request.
export function schedulingText(message: string): string {
  let contentField = false;
  return message.split(/[，,。；;\n]/).filter(clause => {
    if (/^\s*(?:主题|标题|名称|说明|描述)\s*(?:是|为|叫|[:：])/.test(clause)) contentField = true;
    else if (/^\s*(?:(?:开始|结束|截止|计划|安排|起止)\s*)?(?:时间|日期|时段)\s*(?:是|为|改为|[:：])/.test(clause)) contentField = false;
    return !contentField;
  }).join('，');
}
