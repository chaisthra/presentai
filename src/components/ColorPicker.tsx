interface ColorPickerProps {
  color: string;
  onChange: (color: string) => void;
}

export function ColorPicker({ color, onChange }: ColorPickerProps) {
  return (
    <div className="flex items-center gap-2">
      <input
        type="color"
        value={color}
        onChange={(e) => onChange(e.target.value)}
        className="h-8 w-16 rounded cursor-pointer"
      />
      <input
        type="text"
        value={color}
        onChange={(e) => onChange(e.target.value)}
        className="flex-grow rounded-md border-gray-300 shadow-sm focus:border-indigo-500 focus:ring-indigo-500 sm:text-sm"
      />
    </div>
  );
}