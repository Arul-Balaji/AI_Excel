from ship_management_model import ShipManagementModel
from sales_forecast_model import SalesForecastModel

class Agent:
    """
    An agent that can interact with spreadsheet models to perform tasks.
    """

    def __init__(self, tools):
        """
        Initialize the Agent with a list of tools.

        Args:
            tools (list): A list of tool objects (e.g., instantiated models)
        """
        self.tools = {tool.name: tool for tool in tools}
        print(f"Agent initialized with tools: {list(self.tools.keys())}")

    def execute_task(self, prompt):
        """
        Execute a task based on a natural language prompt.
        This is a simplified implementation that uses keyword matching.
        """
        print(f"\n> Executing task: '{prompt}'")
        prompt = prompt.lower()

        # Keyword-based tool selection
        if 'ship' in prompt or 'revenue' in prompt:
            tool = self.tools.get('ship_management')
            if not tool:
                return "Error: Ship management tool not available."

            # Sub-task routing for ship management
            if 'list' in prompt or 'get all' in prompt and 'ship type' in prompt:
                return tool.get_all_ship_types()
            elif 'revenue' in prompt:
                # A more advanced agent would parse the ship name from the prompt
                # For this demo, we'll hardcode one
                ship_name = 'Container Ships'
                print(f"   (Assuming ship name: '{ship_name}')")
                return tool.read_total_revenue(ship_name)
            elif 'add' in prompt and 'ship type' in prompt:
                # A more advanced agent would parse the new ship name
                new_ship_name = 'Bulk Carrier'
                print(f"   (Assuming new ship name: '{new_ship_name}')")
                return tool.add_ship_type(new_ship_name)
            else:
                return "Sorry, I don't understand that ship management command."

        elif 'sales' in prompt or 'forecast' in prompt:
            tool = self.tools.get('sales_forecast')
            if not tool:
                return "Error: Sales forecast tool not available."

            # Sub-task routing for sales forecast
            if 'read' in prompt or 'get' in prompt or 'show' in prompt:
                return tool.read_sales_forecast()
            else:
                return "Sorry, I don't understand that sales forecast command."

        else:
            return "Sorry, I couldn't determine which tool to use for that request."

def get_tools():
    """
    Instantiate and return a list of available tools.
    """
    ship_model = ShipManagementModel()
    sales_model = SalesForecastModel()
    return [ship_model, sales_model]

if __name__ == '__main__':
    # Initialize tools and agent
    tools = get_tools()
    agent = Agent(tools)

    # --- DEMO PROMPTS ---
    print("\n" + "="*80)
    print("DEMONSTRATING AGENT CAPABILITIES")
    print("="*80)

    # Prompt 1: List all ship types
    result1 = agent.execute_task("Can you give me the list of ship types?")
    print("\n< Result for 'list ship types':")
    print(result1)

    # Prompt 2: Get revenue for a specific ship type
    result2 = agent.execute_task("What is the total revenue for Container Ships?")
    print("\n< Result for 'get revenue':")
    print(result2)

    # Prompt 3: Add a new ship type
    # First, restore the original file to have a clean slate
    import shutil
    shutil.copy("Ship Mgt Financial Model v1 - Populated Example.xlsx.bak", "Ship Mgt Financial Model v1 - Populated Example.xlsx")
    result3 = agent.execute_task("Add a new ship type for Bulk Carrier")
    print("\n< Result for 'add ship type':")
    print(result3)

    # Verify that the new ship type was added
    result4 = agent.execute_task("Get all ship types")
    print("\n< Verification after adding:")
    print(result4)
