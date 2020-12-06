/*
 * Copyright (C) 2004-2017 AMain.com, Inc.
 * Copyright 2009-2013 Josh Close
 * All Rights Reserved
 * 
 * See LICENSE.txt for details or visit http://www.opensource.org/licenses/ms-pl.html for MS-PL and http://opensource.org/licenses/Apache-2.0 for Apache 2.0.
 */

using System;
using System.Collections;
using System.Collections.Generic;

namespace ExcelHelper.TypeConversion
{
    /// <summary>
    /// Creates <see cref="ITypeConverter"/>s.
    /// </summary>
    public static class TypeConverterFactory
    {
        private static readonly Dictionary<Type, ITypeConverter> _typeConverters = new Dictionary<Type, ITypeConverter>();
        private static readonly object _locker = new object();

        /// <summary>
        /// Initializes the <see cref="TypeConverterFactory" /> class.
        /// </summary>
        static TypeConverterFactory() 
        {
            AddConverter(typeof(bool), new BooleanConverter());
            AddConverter(typeof(byte), new ByteConverter());
            AddConverter(typeof(char), new CharConverter());
            AddConverter(typeof(DateTime), new DateTimeConverter());
            AddConverter(typeof(decimal), new DecimalConverter());
            AddConverter(typeof(double), new DoubleConverter());
            AddConverter(typeof(float), new SingleConverter());
            AddConverter(typeof(Guid), new GuidConverter());
            AddConverter(typeof(short), new Int16Converter());
            AddConverter(typeof(int), new Int32Converter());
            AddConverter(typeof(long), new Int64Converter());
            AddConverter(typeof(sbyte), new SByteConverter());
            AddConverter(typeof(string), new StringConverter());
            AddConverter(typeof(TimeSpan), new TimeSpanConverter());
            AddConverter(typeof(ushort), new UInt16Converter());
            AddConverter(typeof(uint), new UInt32Converter());
            AddConverter(typeof(ulong), new UInt64Converter());
            AddConverter(typeof(IEnumerable), new EnumerableConverter());
        }

        /// <summary>
        /// Adds the <see cref="ITypeConverter"/> for the given <see cref="Type"/>.
        /// </summary>
        /// <param name="type">The type the converter converts.</param>
        /// <param name="typeConverter">The type converter that converts the type.</param>
        private static void AddConverter(
            Type type,
            ITypeConverter typeConverter)
        {
            lock (_locker) {
                _typeConverters[type] = typeConverter;
            }
        }

        /// <summary>
        /// Gets the converter for the given <see cref="Type"/>.
        /// </summary>
        /// <param name="type">The type to get the converter for.</param>
        /// <returns>The <see cref="ITypeConverter"/> for the given <see cref="Type"/>.</returns>
        public static ITypeConverter GetConverter(
            Type type)
        {
            if (type == null) {
                throw new ArgumentNullException(nameof(type));
            }
            lock (_locker) {
                if (_typeConverters.TryGetValue(type, out var typeConverter)) {
                    return typeConverter;
                }
            }
            if (typeof(IEnumerable).IsAssignableFrom(type)) {
                return GetConverter(typeof(IEnumerable));
            }
            if (typeof(Enum).IsAssignableFrom(type)) {
                AddConverter(type, new EnumConverter(type));
                return GetConverter(type);
            }
            var isGenericType = type.IsGenericType;
            if (isGenericType && type.GetGenericTypeDefinition() == typeof(Nullable<>)) {
                AddConverter(type, new NullableConverter(type));
                return GetConverter(type);
            }
            throw new ExcelTypeConverterException("Unable to convert type '" + type.Name + "'");
        }
    }
}