try:
    from openai.error import RateLimitError, InvalidRequestError
    print("Successfully imported RateLimitError and InvalidRequestError")
except ModuleNotFoundError as e:
    print(f"Error: {e}")