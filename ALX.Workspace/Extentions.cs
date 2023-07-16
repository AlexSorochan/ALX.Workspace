using System;
using System.Linq;
using System.Drawing;
using System.ComponentModel;

namespace ALX.Workspace
{
    /// <summary>
    /// Атрибут цвета
    /// </summary>
    [AttributeUsage(AttributeTargets.Field, Inherited = false, AllowMultiple = true)]
    public sealed class EnumMemberColor : Attribute
    {
        public Color Color { get; }

        public EnumMemberColor(string colorName)
        {
            Color = Color.FromName(colorName);
        }
    }

    public static class Extentions
    {
        /// <summary>
        /// Значение атрибута
        /// </summary>
        /// <typeparam name="TAttribute">Класс атрибута</typeparam>
        /// <param name="value">Член перечисления</param>
        /// <returns></returns>
        public static TAttribute GetAttribute<TAttribute>(this Enum value) where TAttribute : Attribute
        {
            Type type = value.GetType();
            string name = Enum.GetName(type, value);
            return type.GetField(name).GetCustomAttributes(false).
                OfType<TAttribute>().SingleOrDefault();
        }

        /// <summary>
        /// Значение атрибута Description
        /// </summary>
        /// <param name="member">Член перечисления</param>
        /// <returns></returns>
        public static string GetEnumDescription(this Enum member)
        {
            DescriptionAttribute attr = member.GetAttribute<DescriptionAttribute>();
            return attr?.Description ?? member.ToString();
        }
    }
}
