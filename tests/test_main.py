"""Tests for the main module."""


from cornerstone_automation.main import hello_world

class TestHelloWorld:
    """Test cases for the hello_world function."""

    def test_hello_world_default(self) -> None:
        """Test hello_world with default parameter."""
        result = hello_world()
        assert result == "Hello, World!"

    def test_hello_world_with_name(self) -> None:
        """Test hello_world with a specific name."""
        result = hello_world("Alice")
        assert result == "Hello, Alice!"

    def test_hello_world_with_empty_string(self) -> None:
        """Test hello_world with empty string."""
        result = hello_world("")
        assert result == "Hello, !"

    def test_hello_world_with_special_characters(self) -> None:
        """Test hello_world with special characters."""
        result = hello_world("John Doe!")
        assert result == "Hello, John Doe!!"

    def test_hello_world_with_numbers(self) -> None:
        """Test hello_world with numeric string."""
        result = hello_world("123")
        assert result == "Hello, 123!" 