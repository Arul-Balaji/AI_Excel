from agent import Agent, get_tools
import os
import shutil

def main():
    """
    A demonstration script for the agent.
    """
    # --- SETUP ---
    # Create backups of the original spreadsheets to ensure the demo is repeatable
    ship_file = "Ship Mgt Financial Model v1 - Populated Example.xlsx"
    sales_file = "microsoft_Sales forecast tracker small business.xlsx"

    if not os.path.exists(f"{ship_file}.bak"):
        shutil.copy(ship_file, f"{ship_file}.bak")
    else:
        shutil.copy(f"{ship_file}.bak", ship_file)

    if not os.path.exists(f"{sales_file}.bak"):
        shutil.copy(sales_file, f"{sales_file}.bak")
    else:
        shutil.copy(f"{sales_file}.bak", sales_file)

    # --- INITIALIZE AGENT ---
    tools = get_tools()
    agent = Agent(tools)

    # --- DEMO PROMPTS ---
    print("\n" + "="*80)
    print("DEMONSTRATING AGENT CAPABILITIES")
    print("="*80)

    prompts = [
        "List all available ship types.",
        "What is the total revenue for 'Container Ships'?",
        "Add a new ship type named 'Bulk Carrier'.",
        "Show me the sales forecast.",
    ]

    for prompt in prompts:
        result = agent.execute_task(prompt)
        print(f"\n< Result for '{prompt}':")
        print(result)
        print("-" * 80)

if __name__ == '__main__':
    main()
