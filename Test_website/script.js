const mapData = {
  ppi: {
    title: "蛋白互作网络重排",
    body: "从互作证据、蛋白组扰动和结构相似性中识别疾病状态下被重新连接的关键复合体。",
    nodes: ["SRC", "EGFR", "GRB2", "STAT3", "AKT1", "MAPK1"],
    evidence: ["互作置信度 0.86", "扰动一致性高", "文献证据 42 篇"],
  },
  drug: {
    title: "药物响应机制链",
    body: "将药筛曲线、靶点结构和扰动表达签名合并，解释候选药物敏感性与耐药来源。",
    nodes: ["Drug-A", "JAK2", "STAT", "IC50", "Rescue", "Phenotype"],
    evidence: ["药筛 AUC 改善", "靶点口袋稳定", "表达签名回调"],
  },
  receptor: {
    title: "受体功能状态图谱",
    body: "结合配体亲和预测、受体结构口袋和单细胞状态，定位受体激活后的下游信号轴。",
    nodes: ["Ligand", "GPCR", "G-protein", "cAMP", "CREB", "State"],
    evidence: ["结合位点收敛", "细胞状态特异", "通路富集显著"],
  },
};

const positions = [
  [50, 26],
  [28, 48],
  [50, 52],
  [70, 42],
  [36, 72],
  [68, 72],
];

function renderMap(key) {
  const data = mapData[key];
  const title = document.querySelector("[data-map-title]");
  const body = document.querySelector("[data-map-body]");
  const network = document.querySelector("[data-network]");
  const evidence = document.querySelector("[data-evidence]");

  title.textContent = data.title;
  body.textContent = data.body;
  network.innerHTML = `
    <svg viewBox="0 0 100 100" aria-hidden="true">
      <line x1="50" y1="26" x2="28" y2="48"></line>
      <line x1="50" y1="26" x2="70" y2="42"></line>
      <line x1="28" y1="48" x2="50" y2="52"></line>
      <line x1="70" y1="42" x2="50" y2="52"></line>
      <line x1="50" y1="52" x2="36" y2="72"></line>
      <line x1="50" y1="52" x2="68" y2="72"></line>
    </svg>
    ${data.nodes
      .map((node, index) => {
        const [x, y] = positions[index];
        return `<span class="node ${index === 0 ? "core" : ""}" style="left:${x}%;top:${y}%">${node}</span>`;
      })
      .join("")}
  `;
  evidence.innerHTML = data.evidence
    .map((item, index) => `<div><span>证据 ${index + 1}</span><strong>${item}</strong></div>`)
    .join("");
}

document.querySelectorAll("[data-map-key]").forEach((button) => {
  button.addEventListener("click", () => {
    document.querySelectorAll("[data-map-key]").forEach((item) => {
      item.setAttribute("aria-pressed", "false");
    });
    button.setAttribute("aria-pressed", "true");
    renderMap(button.dataset.mapKey);
  });
});

const observer = new IntersectionObserver(
  (entries) => {
    entries.forEach((entry) => {
      if (entry.isIntersecting) {
        entry.target.classList.add("visible");
        observer.unobserve(entry.target);
      }
    });
  },
  { threshold: 0.16 },
);

document.querySelectorAll(".reveal").forEach((element) => observer.observe(element));
renderMap("ppi");
