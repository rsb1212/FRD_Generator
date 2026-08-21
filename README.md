# FRD_Generator

""" 

Target uri-https://bl-prod02-opai-productchangeanalyzer-01.openai.azure.com/openai/responses?api-version=2025-04-01-preview
key-


Example: turning FRD requirement text into a draw.io diagram automatically.

The idea in 3 steps:
1. You already have your requirement/change-request text (from your FRD app).
2. You ask the LLM to convert that text into draw.io's XML "recipe" format.
3. You save the result as a .drawio file — done, no manual drawing.

Install first:  pip install openai --break-system-packages
"""

from openai import OpenAI

client = OpenAI(api_key="YOUR_API_KEY")

# This is the fixed wrapper every .drawio file needs.
# You never change this part — only what goes inside <root>...</root>
DRAWIO_TEMPLATE = """<mxfile>
  <diagram name="Process Flow">
    <mxGraphModel dx="800" dy="600" grid="1" gridSize="10" page="1">
      <root>
        <mxCell id="0" />
        <mxCell id="1" parent="0" />
        {boxes_and_arrows}
      </root>
    </mxGraphModel>
  </diagram>
</mxfile>"""

# The instructions that teach the LLM exactly how to write valid draw.io XML.
SYSTEM_PROMPT = """You convert business process descriptions into draw.io diagram XML.

Rules you must follow:
- Output ONLY the <mxCell> elements that go inside <root> — nothing else, no explanation, no markdown fences.
- Each process step is a box: <mxCell id="X" value="Step name" style="rounded=1;fillColor=#dae8fc;strokeColor=#6c8ebf;" vertex="1" parent="1"><mxGeometry x="..." y="..." width="160" height="60" as="geometry" /></mxCell>
- Stack boxes vertically, 140px apart (e.g. y=40, y=180, y=320, ...), all at x=40.
- Connect each box to the next with an arrow: <mxCell id="X" edge="1" parent="1" source="A" target="B"><mxGeometry relative="1" as="geometry" /></mxCell>
- Use decision/branch boxes (style="rhombus;fillColor=#ffe6cc;") for yes/no or approval steps.
- IDs must be unique integers starting from 2 (0 and 1 are reserved).
- Keep box labels short (3-6 words).
"""


def requirement_to_diagram(requirement_text: str, diagram_type: str = "to-be") -> str:
    """
    Takes plain-English requirement text, returns a complete .drawio XML string.
    diagram_type: "as-is" or "to-be" — just changes the framing of the prompt.
    """
    user_prompt = f"""Convert this {diagram_type} process description into a draw.io process flow:

{requirement_text}

Remember: output ONLY the <mxCell> elements, nothing else."""

    response = client.chat.completions.create(
        model="gpt-4o",
        temperature=0.2,  # low temperature = consistent, predictable structure (important for FRDs)
        messages=[
            {"role": "system", "content": SYSTEM_PROMPT},
            {"role": "user", "content": user_prompt},
        ],
    )

    boxes_and_arrows_xml = response.choices[0].message.content.strip()
    full_drawio_file = DRAWIO_TEMPLATE.format(boxes_and_arrows=boxes_and_arrows_xml)
    return full_drawio_file


if __name__ == "__main__":
    # Example: a sample change request description
    sample_requirement = """
    When a policy servicing request is received, the system first validates
    the customer's identity. If validation fails, the request is rejected
    and the customer is notified. If validation succeeds, the request is
    routed to the underwriting team for review. The underwriter either
    approves or rejects the change. If approved, the policy is updated in
    McCamish NGIN and a confirmation is sent to the customer.
    """

    diagram_xml = requirement_to_diagram(sample_requirement, diagram_type="to-be")

    with open("policy_servicing_to-be.drawio", "w") as f:
        f.write(diagram_xml)

    print("Saved: policy_servicing_to-be.drawio — open this file directly in draw.io")
