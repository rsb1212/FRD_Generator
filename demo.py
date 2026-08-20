import os
from ai_module import _llm

SYSTEM_PROMPT = """You convert business process descriptions into draw.io diagram XML representing a swimlane flowchart.

Rules you must follow:

1. Output ONLY the XML elements that go inside the <root> tag — nothing else, no explanation, no markdown fences.
2. Swimlanes (Pools): Create a vertical container (swimlane) for each unique Actor/Department (e.g., Customer, System, Underwriter) using style="swimlane;whiteSpace=wrap;html=1;".
3. Assign an X coordinate range for each swimlane (e.g., Actor 1 at x=0, Actor 2 at x=240, Actor 3 at x=480, width=240).
4. Use appropriate shapes based on the action type:
   - Start/End (Ellipse): style="ellipse;whiteSpace=wrap;html=1;fillColor=#dae8fc;"
   - Process Steps (Rectangle): style="rounded=1;whiteSpace=wrap;html=1;fillColor=#d5e8d4;"
   - Decision Steps (Diamond): style="rhombus;whiteSpace=wrap;html=1;fillColor=#ffe6cc;"
   - Data/System Record (Cylinder): style="shape=cylinder3;whiteSpace=wrap;html=1;boundedLbl=1;backgroundOutline=1;fillColor=#f8cecc;"
   - Document/Report (Document): style="shape=document;whiteSpace=wrap;html=1;boundedLbl=1;fillColor=#e1d5e7;"
5. Place all shapes inside the appropriate Actor's swimlane by setting their X coordinate within that lane.
6. Layout: Increment the Y coordinate for chronological order (e.g., y=60, y=180, y=300).
7. Connect each step to the next with an arrow (edge). Label arrows coming out of decision boxes with "Yes" or "No" as needed.
8. IDs must be unique integers starting from 2 (0 and 1 are reserved).
9. Keep box labels short (3-6 words). """

DRAWIO_TEMPLATE = """<mxfile host="app.diagrams.net" version="24.4.10">
  <diagram id="ProcessFlow" name="Process Flow">
    <mxGraphModel dx="1000" dy="1000" grid="1" gridSize="10" guides="1" tooltips="1" connect="1" arrows="1" fold="1" page="1" pageScale="1" pageWidth="827" pageHeight="1169" math="0" shadow="0">
      <root>
        <mxCell id="0" />
        <mxCell id="1" parent="0" />
{boxes_and_arrows}
      </root>
    </mxGraphModel>
  </diagram>
</mxfile>"""

def requirement_to_diagram(requirement_text: str, diagram_type: str = "to-be") -> str:
    user_prompt = f"""Convert this {diagram_type} process description into a draw.io process flow:

{requirement_text}

Remember: output ONLY the elements, nothing else."""

    # We use the existing project LLM module for generation
    response = _llm(prompt=user_prompt, system=SYSTEM_PROMPT, provider="gemini")
    
    # Strip any markdown formatting in case the LLM ignored instructions
    boxes_and_arrows_xml = response.strip()
    if boxes_and_arrows_xml.startswith("```xml"):
        boxes_and_arrows_xml = boxes_and_arrows_xml[6:]
    elif boxes_and_arrows_xml.startswith("```"):
        boxes_and_arrows_xml = boxes_and_arrows_xml[3:]
    if boxes_and_arrows_xml.endswith("```"):
        boxes_and_arrows_xml = boxes_and_arrows_xml[:-3]
        
    return DRAWIO_TEMPLATE.format(boxes_and_arrows=boxes_and_arrows_xml.strip())

if __name__ == "__main__":
    sample_requirement = """
    When a policy servicing request is received, the system first validates the customer's identity. 
    If validation fails, the request is rejected and the customer is notified. 
    If validation succeeds, the request is routed to the underwriting team for review. 
    The underwriter either approves or rejects the change. 
    If approved, the policy is updated in McCamish NGIN and a confirmation is sent to the customer.
    """
    
    print("Generating diagram...")
    result = requirement_to_diagram(sample_requirement)
    
    output_file = "sample_swimlane.drawio"
    with open(output_file, "w", encoding="utf-8") as f:
        f.write(result)
        
    print(f"\nResulting Draw.io XML has been saved to: {output_file}")
    print("\nPreview of output:")
    print("-" * 40)
    print(result[:500] + "\n... (truncated for preview)")
