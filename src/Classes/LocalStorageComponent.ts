// useLocalStorage.ts
import { useState } from "react";

function LocalStorageComponent<T>(key: string, initialValue: T) {
  const [storedValue, setStoredValue] = useState<T>(() => {
    try {
      const item = window.localStorage.getItem(key);
      return item ? JSON.parse(item) : initialValue;
    } catch (error) {
      console.warn("Error reading localStorage key “" + key + "”:", error);
      return initialValue;
    }
  });

  const setValue = (value: T) => {
    try {
      setStoredValue(value);
      window.localStorage.setItem(key, JSON.stringify(value));
    } catch (error) {
      console.warn("Error setting localStorage key “" + key + "”:", error);
    }
  };

  return [storedValue, setValue] as const;
}

export default LocalStorageComponent;