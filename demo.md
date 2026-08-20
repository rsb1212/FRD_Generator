The instructions that teach the LLM exactly how to write valid draw.io XML.
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
def requirement_to_diagram(requirement_text: str, diagram_type: str = "to-be") -> str: """ Takes plain-English requirement text, returns a complete .drawio XML string. diagram_type: "as-is" or "to-be" — just changes the framing of the prompt. """ user_prompt = f"""Convert this {diagram_type} process description into a draw.io process flow:

{requirement_text}

Remember: output ONLY the elements, nothing else."""

response = client.chat.completions.create(
    model="Gemini-3.1",
    temperature=0.2,  # low temperature = consistent, predictable structure (important for FRDs)
    messages=[
        {"role": "system", "content": SYSTEM_PROMPT},
        {"role": "user", "content": user_prompt},
    ],
)

boxes_and_arrows_xml = response.choices[0].message.content.strip()
full_drawio_file = DRAWIO_TEMPLATE.format(boxes_and_arrows=boxes_and_arrows_xml)
return full_drawio_file
if name == "main": # Example: a sample change request description sample_requirement = """ When a policy servicing request is received, the system first validates the customer's identity. If validation fails, the request is rejected and the customer is notified. If validation succeeds, the request is routed to the underwriting team for review. The underwriter either approves or rejects the change. If approved, the policy is updated in McCamish NGIN and a confirmation is sent to the customer. """