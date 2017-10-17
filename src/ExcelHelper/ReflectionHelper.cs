/*
 * Copyright (C) 2004-2017 AMain.com, Inc.
 * Copyright 2009-2013 Josh Close
 * All Rights Reserved
 * 
 * See LICENSE.txt for details or visit http://www.opensource.org/licenses/ms-pl.html
 * See LICENSE.txt for details or visit http://www.opensource.org/licenses/ms-pl.html
 */

using System;
using System.Linq;
using System.Linq.Expressions;
using System.Reflection;
using ExcelHelper.Configuration;

namespace ExcelHelper
{
    /// <summary>
    /// Common reflection tasks.
    /// </summary>
    internal static class ReflectionHelper
    {
        /// <summary>
        /// Creates an instance of type T.
        /// </summary>
        /// <typeparam name="T">The type of instance to create.</typeparam>
        /// <returns>A new instance of type T.</returns>
        public static T CreateInstance<T>()
        {
            var constructor = Expression.New(typeof(T));
            var compiled = (Func<T>)Expression.Lambda(constructor).Compile();
            return compiled();
        }

        /// <summary>
        /// Creates an instance of the specified type.
        /// </summary>
        /// <param name="type">The type of instance to create.</param>
        /// <param name="args">The constructor arguments.</param>
        /// <returns>A new instance of the specified type.</returns>
        public static object CreateInstance(
            Type type,
            params object[] args)
        {
            var argumentTypes = args.Select(a => a.GetType()).ToArray();
            var argumentExpressions = argumentTypes.Select(( t, i ) => Expression.Parameter(t, "var" + i)).ToArray();
            var constructorInfo = type.GetConstructor(argumentTypes);
            var constructor = Expression.New(constructorInfo, argumentExpressions);
            var compiled = Expression.Lambda(constructor, argumentExpressions).Compile();
            return compiled.DynamicInvoke(args);
        }

        /// <summary>
        /// Gets the constructor <see cref="NewExpression"/> from the give <see cref="Expression"/>.
        /// </summary>
        /// <typeparam name="T">The <see cref="Type"/> of the object that will be constructed.</typeparam>
        /// <param name="expression">The constructor <see cref="Expression"/>.</param>
        /// <returns>A constructor <see cref="NewExpression"/>.</returns>
        /// <exception cref="System.ArgumentException">Not a constructor expression.;expression</exception>
        public static NewExpression GetConstructor<T>(
            Expression<Func<T>> expression)
        {
            var newExpression = expression.Body as NewExpression;
            if (newExpression == null) {
                throw new ArgumentException("Not a constructor expression.", nameof(expression));
            }

            return newExpression;
        }

        /// <summary>
        /// Gets the property from the expression.
        /// </summary>
        /// <typeparam name="TModel">The type of the model.</typeparam>
        /// <param name="expression">The expression.</param>
        /// <returns>The <see cref="PropertyInfo"/> for the expression.</returns>
        public static PropertyInfo GetProperty<TModel>(
            Expression<Func<TModel, object>> expression)
        {
            var member = GetMemberExpression(expression).Member;
            var property = member as PropertyInfo;
            if (property == null) {
                throw new ExcelConfigurationException($"'{member.Name}' is not a property. Did you try to map a field by accident?");
            }

            return property;
        }

        /// <summary>
        /// Gets the member expression.
        /// </summary>
        /// <typeparam name="TModel">The type of the model.</typeparam>
        /// <typeparam name="T"></typeparam>
        /// <param name="expression">The expression.</param>
        /// <returns></returns>
        private static MemberExpression GetMemberExpression<TModel, T>(
            Expression<Func<TModel, T>> expression)
        {
            // This method was taken from FluentNHibernate.Utils.ReflectionHelper.cs and modified.
            // http://fluentnhibernate.org/
            MemberExpression memberExpression = null;
            if (expression.Body.NodeType == ExpressionType.Convert) {
                var body = (UnaryExpression)expression.Body;
                memberExpression = body.Operand as MemberExpression;
            } else if (expression.Body.NodeType == ExpressionType.MemberAccess) {
                memberExpression = expression.Body as MemberExpression;
            }

            if (memberExpression == null) {
                throw new ArgumentException("Not a member access", nameof(expression));
            }

            return memberExpression;
        }
    }
}